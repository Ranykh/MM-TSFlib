"""MoGU textual (LLM) expert for MM-MoGU.

Wraps a frozen language-model encoder (BERT) behind the same probabilistic
expert contract as the numeric experts: given the batch's matched Time-MMD
text, ``forward`` returns ``(mean, sigma^2)`` with shapes
``[B, pred_len, 1]`` — mean prediction and predictive variance — so the MoE
can gate it by inverse variance exactly like a numeric expert.

Only the small MLP heads are trained; the LLM stays frozen (mirroring the
``use_fullmodel=0`` convention of the stock MM-TSFlib pipeline). Set
``llm_random_init=1`` to build a tiny randomly-initialized encoder instead of
downloading pretrained weights — used for offline/CPU toy runs where only the
pipeline (shapes, gating, variance scales) is being validated.
"""

import numpy as np
import torch
import torch.nn as nn
import torch.nn.functional as F
from transformers import BertConfig, BertModel, BertTokenizer


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
        if llm_name != 'BERT':
            raise ValueError(
                f"TextualExpert currently supports BERT only, got '{llm_name}'")

        if getattr(configs, 'llm_random_init', 0):
            # Tiny randomly-initialized encoder for offline/CPU pipeline tests.
            bert_config = BertConfig(hidden_size=64, num_hidden_layers=2,
                                     num_attention_heads=2, intermediate_size=128)
            self.llm_model = BertModel(bert_config)
            self.tokenizer = None  # basic whitespace hashing tokenizer fallback
            self.d_llm = bert_config.hidden_size
            self.vocab_size = bert_config.vocab_size
        else:
            bert_config = BertConfig.from_pretrained('google-bert/bert-base-uncased')
            bert_config.num_hidden_layers = configs.llm_layers
            try:
                self.llm_model = BertModel.from_pretrained(
                    'google-bert/bert-base-uncased', trust_remote_code=True,
                    local_files_only=True, config=bert_config)
            except EnvironmentError:
                print("Local model files not found. Attempting to download...")
                self.llm_model = BertModel.from_pretrained(
                    'google-bert/bert-base-uncased', trust_remote_code=True,
                    local_files_only=False, config=bert_config)
            try:
                self.tokenizer = BertTokenizer.from_pretrained(
                    'google-bert/bert-base-uncased', trust_remote_code=True,
                    local_files_only=True)
            except EnvironmentError:
                print("Local tokenizer files not found. Attempting to download...")
                self.tokenizer = BertTokenizer.from_pretrained(
                    'google-bert/bert-base-uncased', trust_remote_code=True,
                    local_files_only=False)
            self.d_llm = bert_config.hidden_size
            self.vocab_size = bert_config.vocab_size

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
