from __future__ import annotations

import argparse
import math
import multiprocessing
from pathlib import Path
from typing import Iterable, List, Tuple


def _require_packages() -> Tuple[object, object, object, object, object]:
    try:
        import numpy as np
        import pandas as pd
        import cvxpy as cp
        from joblib import Parallel, delayed
    except ImportError as exc:  # pragma: no cover - runtime check
        raise SystemExit(
            "缺少依赖包，请先安装 numpy、pandas、cvxpy、joblib。"
        ) from exc
    return np, pd, cp, Parallel, delayed


def load_returns(csv_path: Path, pd_module) -> Tuple[List[str], object, object, List[str]]:
    data = pd_module.read_csv(csv_path)
    dates = data["date"].tolist()
    index_returns = data["index"].to_numpy()
    stock_returns = data.drop(columns=["date", "index"]).to_numpy()
    tickers = [c for c in data.columns if c not in {"date", "index"}]
    return dates, index_returns, stock_returns, tickers


def compute_covariance(stock_returns, np_module):
    return np_module.cov(stock_returns, rowvar=False, ddof=1)


def compute_beta(stock_returns, index_returns, np_module):
    cov = np_module.cov(stock_returns.T, index_returns, ddof=1)
    cov_with_index = cov[:-1, -1]
    return cov_with_index / np_module.var(index_returns, ddof=1)


def tracking_error_variance(stock_returns, index_returns, weights, np_module):
    diff = stock_returns @ weights - index_returns
    return float(np_module.var(diff, ddof=1))


def log_likelihood_from_variance(variance: float, sample_size: int) -> float:
    if variance <= 0:
        return float("nan")
    return -0.5 * sample_size * (math.log(2 * math.pi * variance) + 1.0)


def entropy_penalty(weights, cp_module):
    return cp_module.sum(cp_module.kl_div(weights, 1.0)) + cp_module.sum(weights)


def solve_relax_l1_step1(stock_returns, index_returns, lambda1, lambda2, cp_module):
    sample_size, num_assets = stock_returns.shape
    weights = cp_module.Variable(num_assets)
    tracking_error = stock_returns @ weights - index_returns
    objective = (
        cp_module.sum_squares(tracking_error) / sample_size
        + lambda1 * entropy_penalty(weights, cp_module)
        + lambda2 * cp_module.norm1(weights)
    )
    constraints = [weights >= 1e-8]
    problem = cp_module.Problem(cp_module.Minimize(objective), constraints)
    problem.solve(solver=cp_module.SCS, verbose=False)
    return weights.value


def select_assets(weights, threshold: float) -> List[int]:
    if weights is None:
        return [0]
    selected = [i for i, w in enumerate(weights) if w is not None and w > threshold]
    if not selected:
        sorted_idx = sorted(range(len(weights)), key=lambda i: weights[i] or 0.0)
        selected = sorted_idx[-1:]
    return selected


def solve_relax_l1_step2(
    stock_returns,
    index_returns,
    expected_returns,
    mu0,
    lambda1,
    cp_module,
):
    sample_size, num_assets = stock_returns.shape
    weights = cp_module.Variable(num_assets)
    tracking_error = stock_returns @ weights - index_returns
    objective = (
        cp_module.sum_squares(tracking_error) / sample_size
        + lambda1 * entropy_penalty(weights, cp_module)
    )
    constraints = [
        cp_module.sum(weights) == 1.0,
        expected_returns @ weights == mu0,
        weights >= 0.0,
    ]
    problem = cp_module.Problem(cp_module.Minimize(objective), constraints)
    problem.solve(solver=cp_module.SCS, verbose=False)
    return weights.value


def evaluate_lambda2(
    lambda2_value: float,
    stock_returns,
    index_returns,
    expected_returns,
    lambda1,
    mu0,
    threshold,
    cp_module,
    np_module,
):
    first_stage_weights = solve_relax_l1_step1(
        stock_returns, index_returns, lambda1, lambda2_value, cp_module
    )
    selected = select_assets(first_stage_weights, threshold)
    reduced_returns = stock_returns[:, selected]
    reduced_mu = expected_returns[selected]
    second_stage_weights = solve_relax_l1_step2(
        reduced_returns, index_returns, reduced_mu, mu0, lambda1, cp_module
    )
    sample_size = stock_returns.shape[0]
    if second_stage_weights is None:
        variance = float("nan")
        log_likelihood = float("nan")
    else:
        variance = tracking_error_variance(
            reduced_returns, index_returns, second_stage_weights, np_module
        )
        log_likelihood = log_likelihood_from_variance(variance, sample_size)
    return {
        "lambda2": lambda2_value,
        "n": len(selected),
        "variance": variance,
        "log_likelihood": log_likelihood,
    }


def write_results_csv(results: Iterable[dict], output_path: Path, pd_module):
    df = pd_module.DataFrame(results)
    df.to_csv(output_path, index=False)


def write_results_latex(results: Iterable[dict], output_path: Path):
    lines = [
        "\\begin{tabular}{cccc}",
        "\\toprule",
        "$\\lambda_2$ & $n$ & $\\mathrm{Var}$ & $L$ \\\\",
        "\\midrule",
    ]
    for row in results:
        lines.append(
            f"{row['lambda2']:.3f} & {row['n']} & {row['variance']:.4f} & {row['log_likelihood']:.2f} \\\\"
        )
    lines.extend(["\\bottomrule", "\\end{tabular}"])
    output_path.write_text("\n".join(lines), encoding="utf-8")


def parse_lambda2_grid(text: str) -> List[float]:
    return [float(item.strip()) for item in text.split(",") if item.strip()]


def main() -> None:
    np_module, pd_module, cp_module, Parallel, delayed = _require_packages()
    parser = argparse.ArgumentParser(description="Relax-L1两步法指数跟踪优化")
    parser.add_argument(
        "--data",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "data" / "hs300_returns.csv",
    )
    parser.add_argument("--lambda1", type=float, default=0.01)
    parser.add_argument("--lambda2-grid", type=str, default="0.0,0.002,0.004,0.006,0.008,0.010")
    parser.add_argument("--mu0", type=float, default=None)
    parser.add_argument("--threshold", type=float, default=1e-6)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "relax_l1_results.csv",
    )
    parser.add_argument(
        "--latex-table",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "relax_l1_results_table.tex",
    )
    parser.add_argument(
        "--cov-output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "covariance.csv",
    )
    parser.add_argument(
        "--beta-output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "beta.csv",
    )
    parser.add_argument("--jobs", type=int, default=None)
    args = parser.parse_args()

    _, index_returns, stock_returns, tickers = load_returns(args.data, pd_module)
    covariance = compute_covariance(stock_returns, np_module)
    beta = compute_beta(stock_returns, index_returns, np_module)
    expected_returns = stock_returns.mean(axis=0)
    mu0 = args.mu0 if args.mu0 is not None else float(index_returns.mean())
    lambda2_values = parse_lambda2_grid(args.lambda2_grid)

    cpu_total = multiprocessing.cpu_count()
    jobs = args.jobs if args.jobs is not None else max(1, cpu_total - 1)

    results = Parallel(n_jobs=jobs)(
        delayed(evaluate_lambda2)(
            value,
            stock_returns,
            index_returns,
            expected_returns,
            args.lambda1,
            mu0,
            args.threshold,
            cp_module,
            np_module,
        )
        for value in lambda2_values
    )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    if args.cov_output:
        cov_df = pd_module.DataFrame(covariance, columns=tickers, index=tickers)
        cov_df.to_csv(args.cov_output)
    if args.beta_output:
        beta_df = pd_module.DataFrame({"ticker": tickers, "beta": beta})
        beta_df.to_csv(args.beta_output, index=False)
    write_results_csv(results, args.output, pd_module)
    write_results_latex(results, args.latex_table)


if __name__ == "__main__":
    main()
