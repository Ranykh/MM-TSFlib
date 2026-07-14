"""MoGU textual (LLM) expert for MM-MoGU.

Wraps a frozen language-model encoder behind the same probabilistic expert
contract as the numeric experts: given the batch's matched Time-MMD text,
``forward`` returns ``(mean, sigma^2)`` with shapes ``[B, pred_len, 1]`` —
mean prediction and predictive variance — so the MoE can gate it by inverse
variance exactly like a numeric expert.

Supported encoders (``--textual_experts`` names, mirroring the stock
MM-TSFlib ``--llm_model`` zoo): ``BERT``, ``GPT2``, ``GPT2M``, ``GPT2L``,
``GPT2XL`` (public HuggingFace models) and ``LLAMA3`` (gated — requires
``--huggingface_token``). Only the small MLP heads are trained; the LLM stays
frozen (mirroring the ``use_fullmodel=0`` convention of the stock pipeline).

Set ``llm_random_init=1`` to build a tiny randomly-initialized encoder of the
matching family instead of downloading pretrained weights — used for
offline/CPU toy runs where only the pipeline (shapes, gating, variance
scales) is being validated.
"""

import numpy as np
import torch
import torch.nn as nn
import torch.nn.functional as F
from transformers import (
    AutoTokenizer,
    BertConfig, BertModel, BertTokenizer,
    GPT2Config, GPT2Model, GPT2Tokenizer,
    LlamaConfig, LlamaModel,
)

# family -> (config cls, model cls, tokenizer cls); hf_name per expert name
LLM_SPECS = {
    'BERT':   dict(family='bert',  hf_name='google-bert/bert-base-uncased'),
    'GPT2':   dict(family='gpt2',  hf_name='openai-community/gpt2'),
    'GPT2M':  dict(family='gpt2',  hf_name='openai-community/gpt2-medium'),
    'GPT2L':  dict(family='gpt2',  hf_name='openai-community/gpt2-large'),
    'GPT2XL': dict(family='gpt2',  hf_name='openai-community/gpt2-xl'),
    'LLAMA3': dict(family='llama', hf_name='meta-llama/Meta-Llama-3-8B-Instruct'),
}
FAMILY_CLASSES = {
    'bert':  (BertConfig, BertModel, BertTokenizer),
    'gpt2':  (GPT2Config, GPT2Model, GPT2Tokenizer),
    'llama': (LlamaConfig, LlamaModel, AutoTokenizer),
}


class MLP(nn.Module):
    def __init__(self, layer_sizes, dropout_rate=0.3):
        super().__init__()
        self.layers = nn.ModuleList()
        self.dropout = nn.Dropout(dropout_rate)
        for i in range(len(layer_sizes) - 1):
            self.layers.append(nn.Linear(layer_sizes[i], layer_sizes[i + 1]))

    def forward(self, x):
        for i, layer in enumerate(self.layers):
            x = layer(x)
            if i < len(self.layers) - 1:
                x = F.relu(x)
                x = self.dropout(x)
        return x


class Model(nn.Module):
    def __init__(self, configs, llm_name=None):
        super().__init__()
        self.pred_len = configs.pred_len
        self.use_fullmodel = getattr(configs, 'use_fullmodel', 0)
        llm_name = llm_name or getattr(configs, 'llm_model', 'BERT')
        if llm_name not in LLM_SPECS:
            raise ValueError(
                f"TextualExpert supports {sorted(LLM_SPECS)}, got '{llm_name}'")
        spec = LLM_SPECS[llm_name]
        config_cls, model_cls, tokenizer_cls = FAMILY_CLASSES[spec['family']]
        hug_token = getattr(configs, 'huggingface_token', None)

        if getattr(configs, 'llm_random_init', 0):
            # Tiny randomly-initialized encoder of the matching family, for
            # offline/CPU pipeline tests (no download, no token).
            if spec['family'] == 'bert':
                llm_config = BertConfig(hidden_size=64, num_hidden_layers=2,
                                        num_attention_heads=2, intermediate_size=128)
            elif spec['family'] == 'gpt2':
                llm_config = GPT2Config(n_embd=64, n_layer=2, n_head=2)
            else:  # llama
                llm_config = LlamaConfig(hidden_size=64, num_hidden_layers=2,
                                         num_attention_heads=2, num_key_value_heads=2,
                                         intermediate_size=128)
            self.llm_model = model_cls(llm_config)
            self.tokenizer = None  # whitespace-hash tokenizer fallback below
        else:
            load_kwargs = {'trust_remote_code': True}
            if spec['family'] == 'llama':
                # gated repo: needs a HuggingFace token (--huggingface_token)
                load_kwargs['token'] = hug_token
            llm_config = config_cls.from_pretrained(spec['hf_name'], **load_kwargs)
            llm_config.num_hidden_layers = configs.llm_layers  # works for GPT2 too
            llm_config.output_attentions = True                # (attribute_map alias)
            llm_config.output_hidden_states = True
            try:
                self.llm_model = model_cls.from_pretrained(
                    spec['hf_name'], local_files_only=True,
                    config=llm_config, **load_kwargs)
            except EnvironmentError:
                print(f"Local model files for {llm_name} not found. Attempting to download...")
                self.llm_model = model_cls.from_pretrained(
                    spec['hf_name'], local_files_only=False,
                    config=llm_config, **load_kwargs)
            try:
                self.tokenizer = tokenizer_cls.from_pretrained(
                    spec['hf_name'], local_files_only=True, **load_kwargs)
            except EnvironmentError:
                print(f"Local tokenizer for {llm_name} not found. Attempting to download...")
                self.tokenizer = tokenizer_cls.from_pretrained(
                    spec['hf_name'], local_files_only=False, **load_kwargs)
            # GPT2/Llama have no pad token by default: reuse EOS (stock convention)
            if self.tokenizer.pad_token is None:
                if self.tokenizer.eos_token:
                    self.tokenizer.pad_token = self.tokenizer.eos_token
                else:
                    self.tokenizer.add_special_tokens({'pad_token': '[PAD]'})

        self.d_llm = llm_config.hidden_size      # attribute_map covers GPT2's n_embd
        self.vocab_size = llm_config.vocab_size

        # The LLM is frozen; only the heads below train.
        for param in self.llm_model.parameters():
            param.requires_grad = False

        head_sizes = [self.d_llm, max(self.d_llm // 8, 16), self.pred_len]
        self.mean_head = MLP(head_sizes)
        self.unc_head = MLP(head_sizes)

    def _to_prompts(self, batch_text):
        prompts = []
        for t in batch_text:
            if isinstance(t, (list, np.ndarray)):
                t = t[0]
            prompts.append(
                f"<|start_prompt|>Make predictions about the future based on the "
                f"following information: {t}<|end_prompt|>")
        return prompts

    def _tokenize(self, prompts, device):
        if self.tokenizer is not None:
            ids = self.tokenizer(prompts, return_tensors="pt", padding=True,
                                 truncation=True, max_length=1024).input_ids
        else:
            # Random-init mode: hash whitespace tokens into the vocab range.
            max_len = max(len(p.split()) for p in prompts)
            ids = torch.zeros(len(prompts), max_len, dtype=torch.long)
            for b, p in enumerate(prompts):
                for j, w in enumerate(p.split()):
                    ids[b, j] = hash(w) % self.vocab_size
        return ids.to(device)

    def encode(self, batch_text):
        """Encode raw batch text into a pooled embedding of shape [B, d_llm]."""
        device = next(self.parameters()).device
        ids = self._tokenize(self._to_prompts(batch_text), device)
        embeddings = self.llm_model.get_input_embeddings()(ids)  # (B, T, d_llm)
        if self.use_fullmodel:
            embeddings = self.llm_model(inputs_embeds=embeddings).last_hidden_state
        pooled = F.adaptive_avg_pool1d(embeddings.transpose(1, 2), 1).squeeze(2)
        return pooled

    def forward(self, batch_text):
        """Probabilistic expert forward: text -> (mean, sigma^2), each [B, pred_len, 1]."""
        pooled = self.encode(batch_text)
        mean = self.mean_head(pooled).unsqueeze(-1)                       # (B, pred_len, 1)
        sq_sigma = F.softplus(self.unc_head(pooled), threshold=20).unsqueeze(-1)
        return mean, sq_sigma
