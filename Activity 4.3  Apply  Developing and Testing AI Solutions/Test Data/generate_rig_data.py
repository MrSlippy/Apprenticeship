"""
Synthetic gearbox test rig data generator.

Simulates test runs on a gearbox test rig. Each unit is tested for FULL_TEST_MIN
minutes; features are aggregated from the first WINDOW_MIN minutes only, because
the business goal is to predict the final outcome early enough to abort the test.

Three failure modes are simulated so the classes are not linearly separable:
  bearing  -> vibration led
  seal     -> temperature led
  assembly -> torque ripple led

Outputs:
  rig_timeseries.csv  raw per-minute readings (for plots / feature engineering)
  rig_features.csv    one row per unit, model ready
"""

import numpy as np
import pandas as pd

RNG = np.random.default_rng(42)

N_UNITS = 900
FULL_TEST_MIN = 60
WINDOW_MIN = 20          # early window used for prediction
N_SEGMENTS = 4           # window split into blocks to capture curve shape
SAMPLE_PER_MIN = 1

FAIL_RATE = 0.09         # ~9% of units fail outright
MARGINAL_RATE = 0.13     # ~13% scrape through close to the limit

TEMP_LIMIT = 96.0        # peak oil temp acceptance limit (deg C)

VARIANTS = {
    # name: (base temp, base vib, base ripple, ratio, n_gears)
    "LMP-6spd":   (72.0, 1.10, 3.2, 3.45, 6),
    "GT3-6spd":   (70.0, 1.05, 3.0, 3.80, 6),
    "Rally-5spd": (74.0, 1.25, 3.6, 4.10, 5),
    "EV-2spd":    (66.0, 0.85, 2.4, 8.20, 2),
}


def assign_condition(n):
    """Assign each unit a health condition and failure mode."""
    r = RNG.random(n)
    cond = np.where(r < FAIL_RATE, "fail",
           np.where(r < FAIL_RATE + MARGINAL_RATE, "marginal", "pass"))
    mode = np.array(["none"] * n, dtype=object)
    bad = cond != "pass"
    mode[bad] = RNG.choice(["bearing", "seal", "assembly"], size=bad.sum(),
                           p=[0.45, 0.35, 0.20])
    return cond, mode


def simulate_unit(uid, cond, mode, variant, calib_days, ambient, operator):
    base_t, base_v, base_r, ratio, n_gears = VARIANTS[variant]
    t = np.arange(FULL_TEST_MIN * SAMPLE_PER_MIN)
    warmup = 1 - np.exp(-t / 12.0)          # thermal warm-up curve

    # severity: 0 healthy, ~0.35 marginal, ~1.0 failing
    sev = {"pass": 0.0, "marginal": RNG.uniform(0.25, 0.45),
           "fail": RNG.uniform(0.75, 1.3)}[cond]

    # rig drift with calibration age, plus ambient effect
    drift = 0.010 * calib_days
    amb = 0.45 * (ambient - 20.0)

    temp_gain, vib_gain, ripple_gain = 0.0, 0.0, 0.0
    if mode == "seal":
        temp_gain, vib_gain, ripple_gain = 20.0, 0.10, 0.3
    elif mode == "bearing":
        temp_gain, vib_gain, ripple_gain = 5.0, 1.35, 0.5
    elif mode == "assembly":
        temp_gain, vib_gain, ripple_gain = 6.0, 0.35, 3.4

    # Each failure mode degrades with a different SHAPE over time, not just a
    # different magnitude. This is what makes the segmented features earn their
    # place: mean/std over the whole window cannot tell these profiles apart.
    x = t / FULL_TEST_MIN
    if mode == "bearing":
        ramp = x ** 2.2                      # slow start, accelerating
    elif mode == "seal":
        ramp = 1 - np.exp(-x / 0.12)         # early step, then plateaus
    elif mode == "assembly":
        ramp = 0.55 + 0.45 * x               # present from the start, near flat
    else:
        ramp = x ** 1.4

    oil_out = (base_t + amb + drift + 20 * warmup
               + sev * temp_gain * ramp
               + RNG.normal(0, 0.55, t.size))
    oil_in = oil_out - RNG.uniform(6, 9) - 2.0 * warmup + RNG.normal(0, 0.35, t.size)

    vib = (base_v + 0.18 * warmup
           + sev * vib_gain * ramp
           + RNG.normal(0, 0.055, t.size)).clip(0.05)

    ripple = (base_r + sev * ripple_gain * ramp
              + RNG.normal(0, 0.22, t.size)).clip(0.1)

    speed = 4000 + 250 * np.sin(t / 6.0) + RNG.normal(0, 22, t.size)
    torque_in = 300 + 22 * np.sin(t / 5.0) + RNG.normal(0, 4.5, t.size)

    eff = (0.968 - 0.0009 * (oil_out - base_t)
           - 0.006 * sev * ramp + RNG.normal(0, 0.0022, t.size)).clip(0.80, 0.995)
    torque_out = torque_in * ratio * eff
    oil_press = (4.2 - 0.010 * (oil_out - base_t)
                 - 0.25 * sev * ramp * (mode == "seal") + RNG.normal(0, 0.08, t.size))
    acoustic = 78 + 7 * (vib - base_v) + RNG.normal(0, 1.1, t.size)

    return pd.DataFrame({
        "unit_id": uid, "minute": t, "variant": variant, "gear_ratio": ratio,
        "n_gears": n_gears, "operator_id": operator, "calib_days": calib_days,
        "ambient_c": ambient, "oil_temp_in_c": oil_in, "oil_temp_out_c": oil_out,
        "oil_pressure_bar": oil_press, "vibration_rms_mm_s": vib,
        "torque_ripple_nm": ripple, "shaft_speed_rpm": speed,
        "torque_in_nm": torque_in, "torque_out_nm": torque_out,
        "efficiency": eff, "acoustic_db": acoustic,
        "condition": cond, "failure_mode": mode,
    })


SENSORS = ["oil_temp_in_c", "oil_temp_out_c", "oil_pressure_bar",
           "vibration_rms_mm_s", "torque_ripple_nm", "shaft_speed_rpm",
           "torque_in_nm", "torque_out_nm", "efficiency", "acoustic_db"]


def aggregate(df_unit):
    """Collapse the early window into one feature row per unit."""
    w = df_unit[df_unit["minute"] < WINDOW_MIN * SAMPLE_PER_MIN]
    row = {
        "unit_id": w["unit_id"].iloc[0],
        "variant": w["variant"].iloc[0],
        "gear_ratio": w["gear_ratio"].iloc[0],
        "n_gears": w["n_gears"].iloc[0],
        "operator_id": w["operator_id"].iloc[0],
        "calib_days": w["calib_days"].iloc[0],
        "ambient_c": round(w["ambient_c"].iloc[0], 1),
    }
    for s in SENSORS:
        v = w[s].to_numpy()
        row[f"{s}__mean"] = v.mean()
        row[f"{s}__std"] = v.std()
        row[f"{s}__min"] = v.min()
        row[f"{s}__max"] = v.max()
        # derived: rate of change per minute (S9 refinement features)
        row[f"{s}__slope"] = np.polyfit(np.arange(v.size), v, 1)[0]
        row[f"{s}__delta"] = v[-3:].mean() - v[:3].mean()
        # segmented window: 4 blocks of 5 minutes, capturing curve SHAPE
        # (mean+std+slope over the whole window are order-blind; these are not)
        for b, seg in enumerate(np.array_split(v, N_SEGMENTS), start=1):
            row[f"{s}__seg{b}_mean"] = seg.mean()
            row[f"{s}__seg{b}_slope"] = np.polyfit(np.arange(seg.size), seg, 1)[0]
        # curvature: is the rise accelerating or flattening off?
        segs = np.array_split(v, N_SEGMENTS)
        row[f"{s}__accel"] = (np.polyfit(np.arange(segs[-1].size), segs[-1], 1)[0]
                              - np.polyfit(np.arange(segs[0].size), segs[0], 1)[0])

    # targets, taken from the FULL test, not the window
    peak_temp = df_unit["oil_temp_out_c"].max()
    row["target_peak_oil_temp_c"] = peak_temp          # regression target
    row["target_efficiency_loss_pct"] = 100 * (1 - df_unit["efficiency"].min())
    row["condition"] = df_unit["condition"].iloc[0]     # classification target
    row["failure_mode"] = df_unit["failure_mode"].iloc[0]
    row["passed"] = int(peak_temp < TEMP_LIMIT)
    return row


def main():
    cond, mode = assign_condition(N_UNITS)
    variants = RNG.choice(list(VARIANTS), size=N_UNITS, p=[0.34, 0.31, 0.20, 0.15])
    calib = RNG.integers(0, 45, N_UNITS)
    ambient = RNG.normal(20.5, 3.2, N_UNITS)
    operators = RNG.choice([f"OP{i:02d}" for i in range(1, 8)], size=N_UNITS)

    frames = [simulate_unit(f"GBX{i:04d}", cond[i], mode[i], variants[i],
                            int(calib[i]), ambient[i], operators[i])
              for i in range(N_UNITS)]
    ts = pd.concat(frames, ignore_index=True)
    feats = pd.DataFrame([aggregate(f) for f in frames])
 
    # realistic missingness: a sensor drops out on a few runs
    for col in ["acoustic_db__mean", "oil_pressure_bar__mean"]:
        idx = RNG.choice(feats.index, size=int(0.03 * len(feats)), replace=False)
        feats.loc[idx, col] = np.nan

    ts.to_csv("U:/Apprenticeship/Activity 4.3  Apply  Developing and Testing AI Solutions/Test Data/rig_timeseries.csv", index=False)
    feats.to_csv("U:/Apprenticeship/Activity 4.3  Apply  Developing and Testing AI Solutions/Test Data/rig_features.csv", index=False)

    print(f"units: {len(feats)}  columns: {feats.shape[1]}")
    print(feats["condition"].value_counts())
    print(feats["failure_mode"].value_counts())
    print("pass rate by acceptance limit:", feats["passed"].mean().round(3))
    print(feats.groupby("condition")["target_peak_oil_temp_c"].mean().round(1))


if __name__ == "__main__":
    main()
