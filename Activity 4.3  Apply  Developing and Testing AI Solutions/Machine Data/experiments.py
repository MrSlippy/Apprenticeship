"""
Experiment reruns for Activity 4.3 (AI4I 2020 predictive maintenance).

Reproduces the experiments listed in Activity_4_3_Experiment_Log.docx using the same pipeline
as AI4I_Predictive_Maintenance_Workbook.ipynb. Run from the Machine Data folder:

    python experiments.py

Writes ai4i_experiment_reruns.csv and ai4i_seed_variation.csv.
"""
import sys
import numpy as np
import pandas as pd
import sklearn
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import precision_score, recall_score, f1_score, average_precision_score
from sklearn.base import clone

RANDOM_SEED = 42
DATA = sys.argv[1] if len(sys.argv) > 1 else "ai4i2020.csv"

machines = pd.read_csv(DATA)
machines["temp_difference_K"] = machines["Process temperature [K]"] - machines["Air temperature [K]"]
machines["power_W"] = machines["Torque [Nm]"] * machines["Rotational speed [rpm]"] * 2 * np.pi / 60
machines["wear_times_torque"] = machines["Tool wear [min]"] * machines["Torque [Nm]"]

target = "Machine failure"
raw = ["Air temperature [K]", "Process temperature [K]", "Rotational speed [rpm]",
       "Torque [Nm]", "Tool wear [min]"]
engineered = raw + ["temp_difference_K", "power_W", "wear_times_torque"]
leakage = ["TWF", "HDF", "PWF", "OSF", "RNF"]
categorical = ["Type"]


def forest(class_weight="balanced_subsample", min_samples_leaf=2):
    return RandomForestClassifier(n_estimators=300, min_samples_leaf=min_samples_leaf,
                                  class_weight=class_weight, random_state=RANDOM_SEED, n_jobs=-1)


def logistic(C=1.0):
    return LogisticRegression(max_iter=5000, class_weight="balanced", C=C, random_state=RANDOM_SEED)


def run(numeric, algorithm, split_seed=RANDOM_SEED):
    X, y = machines[numeric + categorical], machines[target]
    train_x, test_x, train_y, test_y = train_test_split(
        X, y, test_size=0.25, random_state=split_seed, stratify=y)
    pipeline = Pipeline([
        ("preparation", ColumnTransformer([
            ("numbers", StandardScaler(), numeric),
            ("text", OneHotEncoder(handle_unknown="ignore", sparse_output=False), categorical)])),
        ("algorithm", clone(algorithm))])
    pipeline.fit(train_x, train_y)
    probability = pipeline.predict_proba(test_x)[:, 1]
    prediction = (probability >= 0.5).astype(int)
    return {"precision": precision_score(test_y, prediction, zero_division=0),
            "recall": recall_score(test_y, prediction),
            "f1": f1_score(test_y, prediction, zero_division=0),
            "pr_auc": average_precision_score(test_y, probability),
            "caught": int(((prediction == 1) & (test_y == 1)).sum()),
            "false_alarms": int(((prediction == 1) & (test_y == 0)).sum())}


experiments = [
    ("E1 baseline", "Random forest, raw features", raw, forest()),
    ("E1 baseline", "Logistic regression, raw features", raw, logistic()),
    ("E2 leakage", "Random forest, raw + failure-mode flags", raw + leakage, forest()),
    ("E3 class weighting", "Random forest, raw, class_weight=None", raw, forest(class_weight=None)),
    ("E4 tree depth", "Random forest, raw, min_samples_leaf=1", raw, forest(min_samples_leaf=1)),
    ("E4 tree depth", "Random forest, raw, min_samples_leaf=50", raw, forest(min_samples_leaf=50)),
    ("E5 regularisation", "Logistic regression, raw, C=0.001", raw, logistic(C=0.001)),
    ("E5 regularisation", "Logistic regression, raw, C=1000", raw, logistic(C=1000)),
    ("E6 feature engineering", "Random forest, raw + calculated features", engineered, forest()),
]
rows = [{"experiment": e, "configuration": c, **run(cols, alg)} for e, c, cols, alg in experiments]
reruns = pd.DataFrame(rows).round(3)
reruns.to_csv("ai4i_experiment_reruns.csv", index=False)
print(f"scikit-learn {sklearn.__version__}\n")
print(reruns.to_string(index=False))

# E7: ten random splits (split seeds 0-9), both feature sets on the same splits
seeds = []
for seed in range(10):
    r = run(raw, forest(), split_seed=seed)
    e = run(engineered, forest(), split_seed=seed)
    seeds.append({"split_seed": seed,
                  "raw_pr_auc": r["pr_auc"], "raw_recall": r["recall"], "raw_precision": r["precision"],
                  "engineered_pr_auc": e["pr_auc"], "engineered_recall": e["recall"],
                  "engineered_precision": e["precision"]})
seed_table = pd.DataFrame(seeds).round(3)
seed_table.to_csv("ai4i_seed_variation.csv", index=False)
print("\n", seed_table.to_string(index=False))
summary = seed_table.drop(columns="split_seed").agg(["mean", "std", "min", "max"]).round(3)
print("\n", summary.to_string())
gain = seed_table.engineered_pr_auc - seed_table.raw_pr_auc
print(f"\nengineered ahead on {int((gain > 0).sum())} of 10 splits; "
      f"mean gain {gain.mean():.3f}, smallest {gain.min():.3f}")
