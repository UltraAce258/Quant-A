from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Iterable, List, Tuple


def _require_packages() -> Tuple[object, object]:
    try:
        import pandas as pd
        import numpy as np
    except ImportError as exc:  # pragma: no cover - runtime check
        raise SystemExit("缺少依赖包，请先安装 numpy、pandas。") from exc
    return pd, np


def read_results(csv_path: Path, pd_module):
    return pd_module.read_csv(csv_path)


def compute_bic(variance: float, n: int, sample_size: int) -> float:
    if variance <= 0:
        return float("nan")
    return math.log(variance) + n * math.log(sample_size) / sample_size


def attach_bic(df, sample_size: int, np_module):
    df = df.copy()
    df["bic"] = [
        compute_bic(var, int(n), sample_size)
        for var, n in zip(df["variance"], df["n"])
    ]
    return df


def export_bic_table(df, output_path: Path):
    df.to_csv(output_path, index=False)


def write_bic_plot_tex(df, output_path: Path):
    points = "\n".join(
        f"({int(row.n)}, {row.bic:.6f})" for row in df.itertuples()
    )
    tex = "\n".join(
        [
            "\\begin{tikzpicture}",
            "\\begin{axis}[",
            "width=0.9\\linewidth,",
            "height=0.55\\linewidth,",
            "grid=both,",
            "xlabel={$n$},",
            "ylabel={BIC},",
            "line width=1pt,",
            "samples=50,",
            "]",
            "\\addplot+[mark=o, color=blue] coordinates {",
            points,
            "};",
            "\\end{axis}",
            "\\end{tikzpicture}",
        ]
    )
    output_path.write_text(tex, encoding="utf-8")


def main() -> None:
    pd_module, np_module = _require_packages()
    parser = argparse.ArgumentParser(description="BIC准则计算与最优n筛选")
    parser.add_argument(
        "--input",
        type=Path,
        default=Path(__file__).resolve().parents[1]
        / "outputs"
        / "relax_l1_results.csv",
    )
    parser.add_argument(
        "--sample-size",
        type=int,
        default=None,
        help="样本长度T，若为空则从数据文件读取",
    )
    parser.add_argument(
        "--data",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "data" / "hs300_returns.csv",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "bic_results.csv",
    )
    parser.add_argument(
        "--plot-tex",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "outputs" / "bic_curve.tex",
    )
    args = parser.parse_args()

    results = read_results(args.input, pd_module)
    if args.sample_size is None:
        returns = pd_module.read_csv(args.data)
        sample_size = len(returns)
    else:
        sample_size = args.sample_size

    bic_table = attach_bic(results, sample_size, np_module)
    export_bic_table(bic_table, args.output)
    write_bic_plot_tex(bic_table, args.plot_tex)

    best_row = bic_table.loc[bic_table["bic"].idxmin()]
    print(f"最优n: {int(best_row['n'])}, 对应lambda2: {best_row['lambda2']}")


if __name__ == "__main__":
    main()
