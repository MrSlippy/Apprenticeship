"""
Gearbox rig outcome modelling.

Runs the same pipeline over three progressively richer feature sets so the
S9 refinement can be evidenced with before/after metrics:

  A  static + simple window aggregates (mean, std, min, max)   -- order blind
  B  A + slope and delta                                       -- first-order trend
  C  B + segmented block means/slopes + accel                  -- curve shape

Both tasks are run:
  classification  did it fail?           LogisticRegression vs RandomForest
  regression      peak oil temperature   Ridge vs RandomForestRegressor

Outputs CSV result tables and PNG charts into OUTPUT_FOLDER.
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from sklearn.base import clone
from sklearn.model_selection import train_test_split, StratifiedKFold, cross_val_score
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression, Ridge
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.metrics import (classification_report, confusion_matrix, f1_score,
                             precision_score, recall_score, average_precision_score,
                             precision_recall_curve, mean_absolute_error,
                             mean_squared_error, r2_score)

# ---------------------------------------------------------------- settings
OUTPUT_FOLDER = "U:/Apprenticeship/Activity 4.3  Apply  Developing and Testing AI Solutions/Test Data"   # change this to your own folder
RANDOM_SEED = 42                           # fixes every random step, so results repeat
TEST_FRACTION = 0.25                       # a quarter of gearboxes held back for testing

all_gearboxes = pd.read_csv(f"{OUTPUT_FOLDER}/rig_features.csv")

# Which columns are which.
text_columns = ["variant", "operator_id"]
static_number_columns = ["gear_ratio", "n_gears", "calib_days", "ambient_c"]

# Columns that must NEVER be used as features: they either identify the
# gearbox or contain the answer we are trying to predict.
excluded_columns = ["unit_id", "condition", "failure_mode", "passed",
                    "target_peak_oil_temp_c", "target_efficiency_loss_pct"]

# Every sensor summary column has a double underscore, e.g. oil_temp_out_c__mean
sensor_summary_columns = [c for c in all_gearboxes.columns if "__" in c]


def get_feature_columns(set_name):
    """Return the sensor columns belonging to feature set A, B or C.

    A = order-blind aggregates only
    B = A plus first-order trend (slope, delta)
    C = B plus curve shape (segmented blocks, acceleration)
    """
    def suffix(column):
        return column.split("__")[1]        # 'oil_temp_out_c__mean' -> 'mean'

    aggregates = [c for c in sensor_summary_columns
                  if suffix(c) in ("mean", "std", "min", "max")]
    trend = [c for c in sensor_summary_columns
             if suffix(c) in ("slope", "delta")]
    shape = [c for c in sensor_summary_columns
             if suffix(c).startswith("seg") or suffix(c) == "accel"]

    return {"A": aggregates,
            "B": aggregates + trend,
            "C": aggregates + trend + shape}[set_name]


def make_pipeline(number_columns, algorithm):
    """Glue the preparation steps to the algorithm so they travel together.

    Numeric columns: fill any gaps with the median, then standardise.
    Text columns:    one-hot encode into 0/1 columns.

    Keeping these inside one Pipeline means the median and the scaling are
    learned from the TRAINING rows only, then reused on the test rows.
    Doing it by hand risks leaking test information into training.
    """
    preparation = ColumnTransformer([
        ("numbers", Pipeline([
            ("fill_gaps", SimpleImputer(strategy="median")),
            ("standardise", StandardScaler()),
        ]), number_columns),
        ("text", OneHotEncoder(handle_unknown="ignore", sparse_output=False),
         text_columns),
    ])
    # clone() gives each pipeline its own copy of the algorithm.
    # Without it, all pipelines would share one object and overwrite each other.
    return Pipeline([("preparation", preparation), ("algorithm", clone(algorithm))])


# ------------------------------------------------------------------- split
# Pick which gearboxes go in each half. stratify keeps the failure rate the
# same in both, otherwise a random cut could land most failures on one side.
training_rows, testing_rows = train_test_split(
    all_gearboxes.index,
    test_size=TEST_FRACTION,
    random_state=RANDOM_SEED,
    stratify=all_gearboxes["passed"])

training_set = all_gearboxes.loc[training_rows]
testing_set = all_gearboxes.loc[testing_rows]

print(f"training gearboxes: {len(training_set)}   testing gearboxes: {len(testing_set)}")
print(f"failure rate  training {1 - training_set.passed.mean():.3f}"
      f"   testing {1 - testing_set.passed.mean():.3f}\n")

# ---------------------------------------------------------------- baseline
# The dumbest possible model: assume every gearbox passes.
# This is the number every real model has to beat.
always_pass = np.zeros(len(testing_set))
baseline_accuracy = (always_pass == (1 - testing_set.passed)).mean()
print(f"baseline (assume every gearbox passes): accuracy {baseline_accuracy:.3f}, "
      "failures caught 0\n")

# ---------------------------------------------------------- classification
# We flip the label so that 1 = FAILED. Precision and recall are always
# measured against the class we call positive, and failures are what we
# actually want to find.
did_fail_training = 1 - training_set["passed"]
did_fail_testing = 1 - testing_set["passed"]

classifiers = {
    "LogisticRegression": LogisticRegression(
        max_iter=2000,
        class_weight="balanced",     # counteracts the 18% failure rate
        C=1.0,                       # regularisation strength (lower = stronger)
        random_state=RANDOM_SEED),
    "RandomForest": RandomForestClassifier(
        n_estimators=400,            # number of trees
        min_samples_leaf=3,          # smallest allowed leaf; guards against overfitting
        class_weight="balanced_subsample",
        random_state=RANDOM_SEED,
        n_jobs=-1),
}

classification_results = []
trained_models = {}

for set_name in ["A", "B", "C"]:
    sensor_columns = get_feature_columns(set_name)
    number_columns = sensor_columns + static_number_columns
    input_columns = number_columns + text_columns

    training_inputs = training_set[input_columns]
    testing_inputs = testing_set[input_columns]

    for algorithm_name, algorithm in classifiers.items():
        pipeline = make_pipeline(number_columns, algorithm)
        pipeline.fit(training_inputs, did_fail_training)

        # Probability that each test gearbox will fail.
        failure_probability = pipeline.predict_proba(testing_inputs)[:, 1]
        predicted_failure = (failure_probability >= 0.5).astype(int)

        # Cross-validation on the training half only, as a sanity check that
        # the held-out score is not just a lucky split.
        cross_val_scores = cross_val_score(
            pipeline, training_inputs, did_fail_training,
            cv=StratifiedKFold(5, shuffle=True, random_state=RANDOM_SEED),
            scoring="average_precision")

        classification_results.append({
            "feature_set": set_name,
            "n_features": len(number_columns),
            "algorithm": algorithm_name,
            "precision": precision_score(did_fail_testing, predicted_failure,
                                         zero_division=0),
            "recall": recall_score(did_fail_testing, predicted_failure),
            "f1": f1_score(did_fail_testing, predicted_failure),
            "pr_auc": average_precision_score(did_fail_testing, failure_probability),
            "cv_pr_auc_mean": cross_val_scores.mean(),
            "cv_pr_auc_std": cross_val_scores.std(),
        })
        trained_models[(set_name, algorithm_name)] = (
            pipeline, failure_probability, number_columns)

classification_table = pd.DataFrame(classification_results).round(3)
classification_table.to_csv(f"{OUTPUT_FOLDER}/results_classification.csv", index=False)
print("CLASSIFICATION  (positive class = FAILED)")
print(classification_table.to_string(index=False), "\n")

# -------------------------------------------------------------- regression
peak_temp_training = training_set["target_peak_oil_temp_c"]
peak_temp_testing = testing_set["target_peak_oil_temp_c"]

regressors = {
    "Ridge": Ridge(alpha=1.0, random_state=RANDOM_SEED),
    "RandomForest": RandomForestRegressor(n_estimators=400, min_samples_leaf=3,
                                          random_state=RANDOM_SEED, n_jobs=-1),
}

regression_results = []
for set_name in ["A", "B", "C"]:
    sensor_columns = get_feature_columns(set_name)
    number_columns = sensor_columns + static_number_columns
    input_columns = number_columns + text_columns

    for algorithm_name, algorithm in regressors.items():
        pipeline = make_pipeline(number_columns, algorithm)
        pipeline.fit(training_set[input_columns], peak_temp_training)
        predicted_temp = pipeline.predict(testing_set[input_columns])

        regression_results.append({
            "feature_set": set_name,
            "n_features": len(number_columns),
            "algorithm": algorithm_name,
            "mae": mean_absolute_error(peak_temp_testing, predicted_temp),
            "rmse": float(np.sqrt(mean_squared_error(peak_temp_testing, predicted_temp))),
            "r2": r2_score(peak_temp_testing, predicted_temp),
        })

regression_table = pd.DataFrame(regression_results).round(3)
regression_table.to_csv(f"{OUTPUT_FOLDER}/results_regression.csv", index=False)
print("REGRESSION  (target = peak oil temperature, degC)")
print(regression_table.to_string(index=False), "\n")

# ------------------------------------------------------- best model detail
best_row = classification_table.sort_values("pr_auc", ascending=False).iloc[0]
best_pipeline, best_probability, best_number_columns = trained_models[
    (best_row.feature_set, best_row.algorithm)]
print(f"BEST: {best_row.algorithm} on feature set {best_row.feature_set} "
      f"(PR-AUC {best_row.pr_auc})\n")

best_prediction = (best_probability >= 0.5).astype(int)
matrix = confusion_matrix(did_fail_testing, best_prediction)
print("confusion matrix")
print(pd.DataFrame(matrix,
                   index=["actually PASSED", "actually FAILED"],
                   columns=["predicted PASS", "predicted FAIL"]).to_string(), "\n")
print(classification_report(did_fail_testing, best_prediction,
                            target_names=["PASSED", "FAILED"], digits=3))

pd.DataFrame(matrix,
             index=["actually PASSED", "actually FAILED"],
             columns=["predicted PASS", "predicted FAIL"]
             ).to_csv(f"{OUTPUT_FOLDER}/confusion_matrix.csv")

# --------------------------------------------------------- threshold sweep
# Same model, same predictions. Only the cut-off moves. This is the
# decision-boundary argument for section 4.3 of the report.
sweep_results = []
for threshold in np.arange(0.05, 0.96, 0.05):
    prediction_at_threshold = (best_probability >= threshold).astype(int)
    sweep_results.append({
        "threshold": round(threshold, 2),
        "precision": precision_score(did_fail_testing, prediction_at_threshold,
                                     zero_division=0),
        "recall": recall_score(did_fail_testing, prediction_at_threshold),
        "f1": f1_score(did_fail_testing, prediction_at_threshold, zero_division=0),
        # good gearbox wrongly flagged as failing
        "false_positives": int(((prediction_at_threshold == 1) &
                                (did_fail_testing == 0)).sum()),
        # failing gearbox missed, so the rig runs the full test for nothing
        "false_negatives": int(((prediction_at_threshold == 0) &
                                (did_fail_testing == 1)).sum()),
    })

threshold_table = pd.DataFrame(sweep_results).round(3)
threshold_table.to_csv(f"{OUTPUT_FOLDER}/threshold_sweep.csv", index=False)
print("\nTHRESHOLD SWEEP")
print(threshold_table.to_string(index=False))

# ----------------------------------------------------- feature importances
algorithm_object = best_pipeline.named_steps["algorithm"]
encoded_text_names = list(best_pipeline.named_steps["preparation"]
                          .named_transformers_["text"]
                          .get_feature_names_out(text_columns))
all_feature_names = best_number_columns + encoded_text_names

if hasattr(algorithm_object, "feature_importances_"):
    importances = pd.Series(algorithm_object.feature_importances_,
                            index=all_feature_names)
else:
    importances = pd.Series(np.abs(algorithm_object.coef_[0]),
                            index=all_feature_names)

top_features = importances.sort_values(ascending=False).head(20)
top_features.to_csv(f"{OUTPUT_FOLDER}/feature_importance.csv", header=["importance"])
print("\nTOP 15 FEATURES")
print(top_features.head(15).round(4).to_string())

# ------------------------------------------------------------------ charts
figure, axes = plt.subplots(1, 2, figsize=(12, 4.2))

by_feature_set = classification_table.pivot(
    index="feature_set", columns="algorithm", values="pr_auc")
by_feature_set.plot(kind="bar", ax=axes[0], rot=0)
axes[0].set_title("Classification PR-AUC by feature set")
axes[0].set_ylabel("PR-AUC")
axes[0].set_xlabel("feature set")

precision_points, recall_points, _ = precision_recall_curve(
    did_fail_testing, best_probability)
axes[1].plot(recall_points, precision_points)
axes[1].axhline((did_fail_testing == 1).mean(), ls="--", c="grey",
                label="no-skill baseline")
axes[1].set_xlabel("recall")
axes[1].set_ylabel("precision")
axes[1].set_title(f"Precision-recall: {best_row.algorithm}, set {best_row.feature_set}")
axes[1].legend()

plt.tight_layout()
plt.savefig(f"{OUTPUT_FOLDER}/chart_performance.png", dpi=140)
plt.close()

figure, axis = plt.subplots(figsize=(7, 5))
top_features.head(15)[::-1].plot(kind="barh", ax=axis)
axis.set_title("Top 15 features")
axis.set_xlabel("importance")
plt.tight_layout()
plt.savefig(f"{OUTPUT_FOLDER}/chart_features.png", dpi=140)
plt.close()

# Example temperature curves, one gearbox per failure mode.
timeseries = pd.read_csv(f"{OUTPUT_FOLDER}/rig_timeseries.csv")
figure, axis = plt.subplots(figsize=(8, 4.5))
for mode, label in [("none", "healthy"), ("seal", "seal failure"),
                    ("bearing", "bearing failure"), ("assembly", "assembly failure")]:
    example_unit = timeseries.loc[timeseries.failure_mode == mode, "unit_id"].iloc[0]
    unit_readings = timeseries[timeseries.unit_id == example_unit]
    axis.plot(unit_readings.minute, unit_readings.oil_temp_out_c, label=label)

axis.axvline(20, ls="--", c="grey")
axis.text(20.5, axis.get_ylim()[0] + 1, "prediction window ends",
          fontsize=8, color="grey")
axis.set_xlabel("minute")
axis.set_ylabel("oil temp out (degC)")
axis.set_title("Oil temperature by failure mode")
axis.legend()
plt.tight_layout()
plt.savefig(f"{OUTPUT_FOLDER}/chart_curves.png", dpi=140)
plt.close()

print("\ncharts and tables written to", OUTPUT_FOLDER)