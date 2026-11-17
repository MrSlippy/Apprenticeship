"""
Does the prediction window length change whether trend/shape features earn
their place? At 20 minutes the absolute level already carries the signal.
Earlier in the test, level should be less informative and trajectory more.

Regenerates features at 8, 12, 16 and 20 minutes and compares feature set A
(order-blind aggregates) against set C (aggregates + trend + curve shape).
"""
import numpy as np, pandas as pd
import generate_rig_data as G
from sklearn.base import clone
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import average_precision_score, recall_score, precision_score

SEED = 42
CAT = ["variant", "operator_id"]
STATIC_NUM = ["gear_ratio", "n_gears", "calib_days", "ambient_c"]
RF = RandomForestClassifier(n_estimators=400, min_samples_leaf=3,
                            class_weight="balanced_subsample",
                            random_state=SEED, n_jobs=-1)

# rebuild the same population once, so only the window changes
G.RNG = np.random.default_rng(SEED)
cond, mode = G.assign_condition(G.N_UNITS)
variants = G.RNG.choice(list(G.VARIANTS), size=G.N_UNITS, p=[0.34, 0.31, 0.20, 0.15])
calib = G.RNG.integers(0, 45, G.N_UNITS)
ambient = G.RNG.normal(20.5, 3.2, G.N_UNITS)
ops = G.RNG.choice([f"OP{i:02d}" for i in range(1, 8)], size=G.N_UNITS)
frames = [G.simulate_unit(f"GBX{i:04d}", cond[i], mode[i], variants[i],
                          int(calib[i]), ambient[i], ops[i]) for i in range(G.N_UNITS)]


def run(window):
    G.WINDOW_MIN = window
    feats = pd.DataFrame([G.aggregate(f) for f in frames])
    sensor = [c for c in feats.columns if "__" in c]
    simple = [c for c in sensor if c.split("__")[1] in ("mean", "std", "min", "max")]
    rich = simple + [c for c in sensor if c.split("__")[1] in ("slope", "delta", "accel")
                     or c.split("__")[1].startswith("seg")]
    y = 1 - feats["passed"]
    tr, te = train_test_split(feats.index, test_size=0.25, random_state=SEED, stratify=y)
    out = {}
    for label, cols in [("A", simple), ("C", rich)]:
        num = cols + STATIC_NUM
        pre = ColumnTransformer([
            ("num", Pipeline([("i", SimpleImputer(strategy="median")),
                              ("s", StandardScaler())]), num),
            ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), CAT)])
        pipe = Pipeline([("pre", pre), ("est", clone(RF))])
        pipe.fit(feats.loc[tr, num + CAT], y.loc[tr])
        prob = pipe.predict_proba(feats.loc[te, num + CAT])[:, 1]
        pred = (prob >= 0.5).astype(int)
        out[label] = {"pr_auc": average_precision_score(y.loc[te], prob),
                      "recall": recall_score(y.loc[te], pred),
                      "precision": precision_score(y.loc[te], pred, zero_division=0)}
    return out


rows = []
for w in [8, 12, 16, 20]:
    r = run(w)
    rows.append({"window_min": w,
                 "A_pr_auc": r["A"]["pr_auc"], "C_pr_auc": r["C"]["pr_auc"],
                 "gain": r["C"]["pr_auc"] - r["A"]["pr_auc"],
                 "A_recall": r["A"]["recall"], "C_recall": r["C"]["recall"]})

res = pd.DataFrame(rows).round(3)
res.to_csv("/mnt/user-data/outputs/results_window_length.csv", index=False)
print(res.to_string(index=False))
