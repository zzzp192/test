#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Origin-free curve plotting and insertion into the report templates."""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

from curve_data import (
    chunk_curves,
    dataframe_to_curves,
    load_tensile_xy_dataframe,
    load_vda_xy_dataframe,
)


PLOT_WIDTH_CM = 16.0
PLOT_HEIGHT_CM = 12.0
PLOT_DPI = 300


def _axis_labels(data_type: str, swap_xy: bool) -> tuple[str, str]:
    if data_type == "tensile":
        labels = ("Engineering strain/%", "Engineering stress/MPa")
        return labels if swap_xy else labels[::-1]
    if data_type == "vda":
        labels = ("Displacement/mm", "Force/kN")
        return labels if swap_xy else labels[::-1]
    return "X", "Y"


def _curve_group_label(sample_label: str) -> str:
    """Collapse sample suffixes such as Group-1/2/3 into one legend label."""
    label = str(sample_label).strip()
    match = re.match(r"^(.*)-\d+$", label)
    return match.group(1) if match else label


def create_transparent_curve_images(
    dataframe,
    output_dir: str,
    lines_per_graph: int,
    data_type: str,
    swap_xy: bool,
) -> list[str]:
    """Render fixed-size 16×12 cm transparent PNGs with matplotlib."""
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import pyplot as plt

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    curve_groups = chunk_curves(dataframe_to_curves(dataframe), lines_per_graph)
    x_label, y_label = _axis_labels(data_type, swap_xy)
    image_paths: list[str] = []

    for group_index, curve_group in enumerate(curve_groups, start=1):
        figure, axes = plt.subplots(
            figsize=(PLOT_WIDTH_CM / 2.54, PLOT_HEIGHT_CM / 2.54),
            dpi=PLOT_DPI,
        )
        figure.patch.set_alpha(0)
        axes.patch.set_alpha(0)

        color_cycle = ["#000000", "#ff0000", "#00d000", "#0000ff", "#d000d0", "#00a0a0", "#ff8800", "#7030a0"]
        group_colors: dict[str, str] = {}
        legend_groups: set[str] = set()
        for curve in curve_group:
            group_label = _curve_group_label(curve.label)
            if group_label not in group_colors:
                group_colors[group_label] = color_cycle[len(group_colors) % len(color_cycle)]
            legend_label = group_label if group_label not in legend_groups else "_nolegend_"
            legend_groups.add(group_label)
            axes.plot(
                curve.x,
                curve.y,
                color=group_colors[group_label],
                linewidth=1.0,
                label=legend_label,
            )

        axes.set_xlabel(x_label, fontfamily="Times New Roman", fontsize=11)
        axes.set_ylabel(y_label, fontfamily="Times New Roman", fontsize=11)
        axes.margins(x=0.0, y=0.03)
        axes.minorticks_on()
        axes.tick_params(axis="both", which="major", direction="in", length=6, width=1.0, labelsize=9)
        axes.tick_params(axis="both", which="minor", direction="in", length=3, width=0.8)
        for tick_label in [*axes.get_xticklabels(), *axes.get_yticklabels()]:
            tick_label.set_fontfamily("Times New Roman")
        for spine in axes.spines.values():
            spine.set_color("#000000")
            spine.set_linewidth(1.15)

        if legend_groups:
            legend = axes.legend(
                loc="lower left",
                bbox_to_anchor=(0.14, 0.18),
                fontsize=7.5,
                frameon=False,
                handlelength=3.0,
                handletextpad=0.4,
            )
            for text in legend.get_texts():
                text.set_fontfamily("Microsoft YaHei")
        figure.subplots_adjust(left=0.14, right=0.985, bottom=0.16, top=0.965)

        image_path = output / f"curve_{group_index:03d}.png"
        figure.savefig(
            image_path,
            dpi=PLOT_DPI,
            transparent=True,
            facecolor="none",
            edgecolor="none",
        )
        plt.close(figure)
        image_paths.append(str(image_path))

    return image_paths


def insert_curve_images(
    ppt_path: str,
    image_paths: list[str],
    width_cm: float = PLOT_WIDTH_CM,
    height_cm: float = PLOT_HEIGHT_CM,
) -> int:
    """Place one fixed-size plot on the right side of each report slide."""
    presentation = Presentation(ppt_path)
    if len(image_paths) > len(presentation.slides):
        raise ValueError(
            f"曲线图数量（{len(image_paths)}）多于报告页数（{len(presentation.slides)}）。"
            "请增大“每图曲线数”，确保每页最多放置一张图。"
        )

    width = Cm(width_cm)
    height = Cm(height_cm)
    left = max(0, presentation.slide_width - width)
    top = max(0, (presentation.slide_height - height) // 2)

    for slide_index, image_path in enumerate(image_paths):
        picture = presentation.slides[slide_index].shapes.add_picture(
            image_path,
            left,
            top,
            width=width,
            height=height,
        )
        picture.name = f"Matplotlib 曲线图 {slide_index + 1}"

    presentation.save(ppt_path)
    return len(image_paths)


def _create_one_click_ppt(
    data_path: str,
    template_path: str,
    output_path: str,
    lines_per_graph: int,
    swap_xy: bool,
    data_type: str,
    report_generator,
) -> str:
    loader = load_tensile_xy_dataframe if data_type == "tensile" else load_vda_xy_dataframe
    dataframe = loader(data_path, swap_xy=swap_xy)

    output_folder = os.path.dirname(os.path.abspath(output_path))
    with tempfile.TemporaryDirectory(prefix="yucaitang_plot_", dir=output_folder) as temp_dir:
        image_paths = create_transparent_curve_images(
            dataframe,
            temp_dir,
            lines_per_graph,
            data_type,
            swap_xy,
        )
        report_message = report_generator(data_path, template_path, output_path)
        if not os.path.exists(output_path):
            raise RuntimeError(report_message or "报告生成失败")
        inserted = insert_curve_images(output_path, image_paths)

    return (
        f"成功生成一键 PPT！\n"
        f"已插入 {inserted} 张透明曲线图（12×16 cm，高×宽）。\n"
        f"保存至: {output_path}"
    )


def create_tensile_one_click_ppt(
    data_path: str,
    template_path: str,
    output_path: str,
    lines_per_graph: int = 12,
    swap_xy: bool = True,
    elongation_mode: str = "ag",
) -> str:
    import tensile_processor

    return _create_one_click_ppt(
        data_path,
        template_path,
        output_path,
        lines_per_graph,
        swap_xy,
        "tensile",
        lambda source, template, target: tensile_processor.generate_report(
            source,
            template,
            target,
            elongation_mode=elongation_mode,
        ),
    )


def create_vda_one_click_ppt(
    data_path: str,
    template_path: str,
    output_path: str,
    lines_per_graph: int = 12,
    swap_xy: bool = True,
) -> str:
    import vda_processor

    return _create_one_click_ppt(
        data_path,
        template_path,
        output_path,
        lines_per_graph,
        swap_xy,
        "vda",
        vda_processor.process_vda_report,
    )
