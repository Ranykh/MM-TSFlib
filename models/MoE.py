import torch
import torch.nn as nn
import torch.nn.functional as F
from models.PatchTST import Model as GatingModel
import copy

class Gating(nn.Module): 
    def __init__(self, configs, individual=False):
        """
        individual: Bool, whether shared model among different variates.
        """
        super(Gating, self).__init__()
        self.projection = nn.Linear(1, configs.num_experts)
        configs = copy.deepcopy(configs)
        configs.prob_expert = 0
        self.gating_arc = GatingModel(configs)

    def gating(self, x_enc):
        enc_out = self.gating_arc.forecast(x_enc, None, None, None) # batch_size, seq_len, variate_dim
        enc_out = enc_out.unsqueeze(-1)
        weights = self.projection(enc_out)
        return weights # batch_size, seq_len, variate_dim, num_experts

    def forward(self, x_enc):
        weights = self.gating(x_enc)
        weights = F.softmax(weights, dim=-1)
        weights = weights.permute(0, 3, 1, 2)
        return weights



class Model(nn.Module):
    def __init__(self, configs, expert_models, textual_experts=None):
        super(Model, self).__init__()
        # expert_models: either a list of already-constructed numeric expert
        # instances (heterogeneous MM-MoGU path) or a single class to repeat
        # num_experts times (backward-compatible homogeneous path).
        # textual_experts: optional list of already-constructed textual (LLM)
        # expert instances whose forward is (batch_text) -> (mean, sigma^2).
        self.unc_gating = configs.unc_gating
        self.prob_expert = configs.prob_expert
        self.task_name = configs.task_name
        # 'none'         -> raw 1/sigma^2 weighting (same-scale experts)
        # 'per_modality' -> z-score log-variance within each modality before
        #                   softmax, so text and numeric variances become
        #                   comparable and the gate does not collapse.
        self.inv_var_norm = getattr(configs, 'inv_var_norm', 'none')

        if isinstance(expert_models, (list, nn.ModuleList)):
            self.experts = nn.ModuleList([m.float() for m in expert_models])
        else:
            self.experts = nn.ModuleList(
                [expert_models(configs).float() for _ in range(configs.num_experts)]
            )
        self.textual_experts = nn.ModuleList(
            [m.float() for m in (textual_experts or [])])
        self.num_experts = len(self.experts) + len(self.textual_experts)
        # 0 = numeric, 1 = textual; expert order is numeric first, then textual
        self.modality_ids = [0] * len(self.experts) + [1] * len(self.textual_experts)

        if not self.unc_gating:
            self.gating = Gating(configs).float()

        if self.textual_experts and not self.prob_expert:
            raise ValueError("textual experts require prob_expert=1 (they emit (mean, sigma^2))")

    def _inverse_variance_weights(self, expert_unc):
        if self.inv_var_norm == 'per_modality' and len(set(self.modality_ids)) > 1:
            # Standardize log-variance within each modality, then softmax over
            # experts: fixes the text-vs-numeric variance-scale collapse.
            log_v = torch.log(expert_unc + 1e-8)
            conf = torch.zeros_like(log_v)
            ids = torch.tensor(self.modality_ids, device=log_v.device)
            for m in ids.unique():
                sel = ids == m
                sub = log_v[:, sel]
                conf[:, sel] = -(sub - sub.mean()) / (sub.std() + 1e-8)
            return F.softmax(conf, dim=1)
        # Raw inverse variance: higher confidence (lower uncertainty) -> higher weight
        inv_var = 1.0 / (expert_unc + 1e-8)   # [batch_size, num_experts, pred_len, num_features]
        sum_inv_var = torch.sum(inv_var, dim=1, keepdim=True)
        return inv_var / sum_inv_var

    def forward(self, x_enc, x_mark_enc, x_dec, x_mark_dec, mask=None, batch_text=None):
        if self.task_name == 'long_term_forecast':
            expert_out = []
            expert_unc = []
            for expert in self.experts:
                sq_sigma = None
                if self.prob_expert:
                    dec_out, sq_sigma = expert(x_enc, x_mark_enc,
                    x_dec, x_mark_dec, mask=None)
                    expert_unc.append(sq_sigma)
                else:
                    dec_out = expert.forward(x_enc, x_mark_enc,
                        x_dec, x_mark_dec, mask=None)
                expert_out.append(dec_out)

            if self.textual_experts:
                if batch_text is None:
                    raise ValueError("MoE has textual experts but no batch_text was provided")
                for expert in self.textual_experts:
                    dec_out, sq_sigma = expert(batch_text)
                    expert_out.append(dec_out)
                    expert_unc.append(sq_sigma)

            expert_out = torch.stack(expert_out, dim=1) # [batch_size, num_experts, pred_len, num_features]
            if len(expert_unc) > 0:
                expert_unc = torch.stack(expert_unc, dim=1) # [batch_size, num_experts, pred_len, num_features]
            if self.unc_gating:
                weights = self._inverse_variance_weights(expert_unc)
            else:
                weights = self.gating(x_enc) # [batch_size, num_experts, pred_len, num_featuers]
            return expert_out, expert_unc, weights
        else:
            raise NotImplementedError("{} not supported with MoE".format(self.task_name))
            