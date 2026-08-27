#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Shared curve-data loading and XY grouping for Origin and matplotlib."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd


@dataclass(frozen=True)
class CurveSeries:
    """One numeric XY curve ready for plotting."""

    label: str
    x: pd.Series
    y: pd.Series


def _normalize_header(value) -> str:
    return re.sub(r"\s+", "", str(value)).lower()


def get_sample_ids_from_excel(file_path: str) -> list[str]:
    """Read the first non-empty sample-ID column found in an Excel workbook."""
    xls = pd.ExcelFile(file_path)
    try:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for column in df.columns:
                if "试样编号" not in str(column):
                    continue
                sample_ids = [
                    str(value).strip()
                    for value in df[column]
                    if pd.notna(value) and str(value).strip()
                ]
                if sample_ids:
                    return sample_ids
    finally:
        xls.close()
    return []


def get_tensile_sample_ids(file_path: str) -> list[str]:
    """Support both the legacy and the newer tensile-summary layouts."""
    xls = pd.ExcelFile(file_path)
    try:
        for sheet_name in xls.sheet_names:
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            for row_idx in range(min(10, len(df_raw))):
                row_values = [str(value) for value in df_raw.iloc[row_idx] if pd.notna(value)]
                if not any("试样编号" in value for value in row_values):
                    continue
                for col_idx, value in enumerate(df_raw.iloc[row_idx]):
                    if pd.isna(value) or "试样编号" not in str(value):
                        continue
                    sample_ids = [
                        str(item).strip()
                        for item in df_raw.iloc[row_idx + 1 :, col_idx]
                        if pd.notna(item) and str(item).strip()
                    ]
                    if sample_ids:
                        return sample_ids
    finally:
        xls.close()
    return []


def is_tensile_curve_columns(columns: Iterable[object]) -> bool:
    columns = list(columns)
    if len(columns) < 2 or len(columns) % 2 != 0:
        return False
    for idx in range(0, len(columns), 2):
        stress = _normalize_header(columns[idx])
        strain = _normalize_header(columns[idx + 1])
        if not ("应力" in stress or "stress" in stress):
            return False
        if not ("应变" in strain or "strain" in strain):
            return False
    return True


def select_tensile_curve_sheet(file_path_or_excel) -> str:
    xls = file_path_or_excel
    owns_excel_file = False
    if not hasattr(xls, "sheet_names"):
        xls = pd.ExcelFile(file_path_or_excel)
        owns_excel_file = True

    try:
        preferred_sheets = [sheet for sheet in xls.sheet_names if "曲线" in sheet]
        preferred_sheets.extend(
            sheet for sheet in xls.sheet_names
            if "原始数据" in sheet and sheet not in preferred_sheets
        )
        preferred_sheets.extend(
            sheet for sheet in xls.sheet_names if sheet not in preferred_sheets
        )
        for sheet in preferred_sheets:
            headers = pd.read_excel(xls, sheet_name=sheet, nrows=0).columns
            if is_tensile_curve_columns(headers):
                return sheet
    finally:
        if owns_excel_file:
            xls.close()
    raise ValueError("未找到有效的拉伸曲线数据工作表")


def is_vda_curve_columns(columns: Iterable[object]) -> bool:
    columns = list(columns)
    if len(columns) < 2 or len(columns) % 2 != 0:
        return False
    valid_pairs = 0
    for idx in range(0, len(columns), 2):
        force = _normalize_header(columns[idx])
        displacement = _normalize_header(columns[idx + 1])
        force_like = "力" in force or "force" in force
        displacement_like = any(
            token in displacement for token in ("位移", "行程", "挠度", "displacement")
        )
        if force_like and displacement_like:
            valid_pairs += 1
    return valid_pairs == len(columns) // 2


def select_vda_curve_sheet(file_path_or_excel) -> str:
    xls = file_path_or_excel
    owns_excel_file = False
    if not hasattr(xls, "sheet_names"):
        xls = pd.ExcelFile(file_path_or_excel)
        owns_excel_file = True

    try:
        preferred_sheets = [sheet for sheet in xls.sheet_names if "原始数据" in sheet]
        preferred_sheets.extend(
            sheet for sheet in xls.sheet_names
            if "VDA" in sheet and sheet not in preferred_sheets
        )
        preferred_sheets.extend(
            sheet for sheet in xls.sheet_names if sheet not in preferred_sheets
        )
        for sheet in preferred_sheets:
            headers = pd.read_excel(xls, sheet_name=sheet, nrows=0).columns
            if is_vda_curve_columns(headers):
                return sheet
    finally:
        if owns_excel_file:
            xls.close()
    raise ValueError("未找到有效的 VDA 力-位移曲线数据工作表")


def prepare_xy_dataframe(
    source_df: pd.DataFrame,
    sample_ids: Iterable[str],
    swap_xy: bool,
) -> pd.DataFrame:
    """Return strict XYXY columns using the same pair-swap rule as Origin."""
    pair_count = len(source_df.columns) // 2
    if pair_count == 0:
        raise ValueError("曲线数据不足，至少需要一组 XY 列")

    sample_ids = list(sample_ids)
    column_positions: list[int] = []
    output_headers: list[str] = []
    for pair_idx in range(pair_count):
        first_idx = pair_idx * 2
        second_idx = first_idx + 1
        x_idx, y_idx = (second_idx, first_idx) if swap_xy else (first_idx, second_idx)
        column_positions.extend([x_idx, y_idx])
        output_headers.extend([
            str(source_df.columns[x_idx]),
            sample_ids[pair_idx] if pair_idx < len(sample_ids) else str(source_df.columns[y_idx]),
        ])

    result = source_df.iloc[:, column_positions].copy()
    result.columns = output_headers
    return result


def dataframe_to_curves(dataframe: pd.DataFrame) -> list[CurveSeries]:
    curves: list[CurveSeries] = []
    for idx in range(0, len(dataframe.columns) - 1, 2):
        numeric = dataframe.iloc[:, [idx, idx + 1]].apply(pd.to_numeric, errors="coerce").dropna()
        if numeric.empty:
            continue
        curves.append(
            CurveSeries(
                label=str(dataframe.columns[idx + 1]),
                x=numeric.iloc[:, 0],
                y=numeric.iloc[:, 1],
            )
        )
    if not curves:
        raise ValueError("曲线数据中没有可绘制的数值 XY 对")
    return curves


def chunk_curves(curves: Iterable[CurveSeries], lines_per_graph: int) -> list[list[CurveSeries]]:
    if lines_per_graph < 1:
        raise ValueError("每图曲线数必须大于 0")
    curves = list(curves)
    return [curves[idx : idx + lines_per_graph] for idx in range(0, len(curves), lines_per_graph)]


def load_tensile_xy_dataframe(file_path: str, swap_xy: bool = True) -> pd.DataFrame:
    if os.path.splitext(file_path)[1].lower() not in (".xlsx", ".xls"):
        raise ValueError("一键 PPT 的拉伸曲线绘图需要 Excel 原始数据（.xlsx/.xls）")
    sample_ids = get_tensile_sample_ids(file_path)
    xls = pd.ExcelFile(file_path)
    try:
        sheet = select_tensile_curve_sheet(xls)
        source_df = pd.read_excel(xls, sheet_name=sheet)
    finally:
        xls.close()
    return prepare_xy_dataframe(source_df, sample_ids, swap_xy)


def load_vda_xy_dataframe(file_path: str, swap_xy: bool = True) -> pd.DataFrame:
    sample_ids = get_sample_ids_from_excel(file_path)
    xls = pd.ExcelFile(file_path)
    try:
        sheet = select_vda_curve_sheet(xls)
        source_df = pd.read_excel(xls, sheet_name=sheet)
    finally:
        xls.close()
    return prepare_xy_dataframe(source_df, sample_ids, swap_xy)
