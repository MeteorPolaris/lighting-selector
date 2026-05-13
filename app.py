from pathlib import Path
import re

import pandas as pd
import streamlit as st
from jinja2 import Environment, FileSystemLoader
from PIL import Image


EXCEL_DIR = Path("excel")
SHEET_NAME = "Sheet Database"
HEADER_ROW_INDEX = 1
CN_HEADER_ROW_INDEX = 2
DATA_START_ROW_INDEX = 3
TEMPLATE_DIR = Path("templates")
SPEC_TEMPLATE_FILE = "spec_sheet.html"
GENERATED_PDF_DIR = Path(".generated/pdfs")
PRODUCT_IMAGE_DIR_CANDIDATES = ["products", "product", "产品图"]
DIMENSION_IMAGE_DIR_CANDIDATES = ["dimensions", "dimension", "sizes", "尺寸图"]
NORMALIZED_IMAGE_DIR = Path(".generated/normalized_images")
LOGO_IMAGE_FILE = Path("assets") / "logo.png"

CATEGORIES: list[dict[str, object]] = [
    {"name": "室外路灯泛光灯", "file": "产品自动选型工具_1 Leon&Yawen_室外路灯泛光灯.xlsx", "assets": Path("assets"), "image_subdir": "1_室外路灯泛光灯"},
    {"name": "室内办公工业", "file": "产品自动选型工具_2 Claire&Sonny_室内办公工业.xlsx", "assets": Path("assets"), "image_subdir": "2_室内办公工业"},
    {"name": "室内筒射灯灯带", "file": "产品自动选型工具_3 Eric&&Yvonne&Tiffany_室内筒射灯灯带.xlsx", "assets": Path("assets"), "image_subdir": "3_室内筒射灯灯带"},
    {"name": "室外景观亮化", "file": "产品自动选型工具_4 ChenChen _室外景观亮化.xlsx", "assets": Path("assets"), "image_subdir": "4_室外景观亮化"},
    {"name": "DDP", "file": "产品自动选型工具_5 Raina&Ivy_DDP.xlsx", "assets": Path("assets"), "image_subdir": "5_DDP"},
]


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").strip()


def _find_column(columns: list[str], keywords: list[str]) -> str | None:
    for col in columns:
        col_lower = col.lower()
        if any(keyword.lower() in col_lower for keyword in keywords):
            return col
    return None


def _parse_code_token(code_token: object) -> list[int]:
    if pd.isna(code_token):
        return []

    text = str(code_token).strip()
    if not text:
        return []

    # 支持 "1~3" / "1-3" / "35" 这类写法
    range_match = re.match(r"^\s*(\d+)\s*[~-]\s*(\d+)\s*$", text)
    if range_match:
        start = int(range_match.group(1))
        end = int(range_match.group(2))
        if start <= end:
            return list(range(start, end + 1))
        return list(range(end, start + 1))

    single_match = re.match(r"^\s*(\d+)\s*$", text)
    if single_match:
        return [int(single_match.group(1))]

    return []


def _extract_code_from_filename(file_path: Path) -> int | None:
    match = re.search(r"\d+", file_path.stem)
    return int(match.group(0)) if match else None


def _extract_codes_from_filename(file_path: Path) -> list[int]:
    stem = file_path.stem.strip()
    parsed = _parse_code_token(stem)
    if parsed:
        return parsed

    single = _extract_code_from_filename(file_path)
    return [single] if single is not None else []


def _resolve_candidate_dir(root_dir: Path, candidates: list[str]) -> Path | None:
    for name in candidates:
        candidate = root_dir / name
        if candidate.exists() and candidate.is_dir():
            return candidate
    return None


def _normalize_image_to_jpg(src_path: Path, category: str) -> Path | None:
    if not src_path.exists():
        return None
    category_dir = NORMALIZED_IMAGE_DIR / category
    category_dir.mkdir(parents=True, exist_ok=True)
    dst_path = category_dir / f"{src_path.stem}.jpg"
    try:
        with Image.open(src_path) as img:
            if img.mode in ("RGBA", "LA", "P"):
                rgb_image = Image.new("RGB", img.size, (255, 255, 255))
                rgb_image.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[-1])
                rgb_image.save(dst_path, "JPEG", quality=95)
            else:
                img.convert("RGB").save(dst_path, "JPEG", quality=95)
        return dst_path
    except Exception:
        return None


def _to_display_text(value: object) -> str:
    if pd.isna(value):
        return "/"
    text = str(value).strip()
    return text if text else "/"


def _to_code_int(value: object) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    # 兼容 1 / 1.0 / "001"
    try:
        return int(float(text))
    except ValueError:
        match = re.search(r"\d+", text)
        return int(match.group(0)) if match else None


def _find_value_from_row(row: pd.Series, keywords: list[str]) -> str:
    for col in row.index:
        col_lower = str(col).lower()
        if any(k.lower() in col_lower for k in keywords):
            return _to_display_text(row[col])
    return "/"


def _build_spec_sections(row: pd.Series) -> list[dict[str, object]]:
    return [
        {
            "en": "Information",
            "cn": "基本信息",
            "items": [
                {"en": "MODEL", "cn": "灯具型号", "value": _find_value_from_row(row, ["型号", "model"])},
                {"en": "TYPE", "cn": "灯具类型", "value": _find_value_from_row(row, ["灯具类型", "type"])},
                {"en": "FAMILY", "cn": "产品系列", "value": _find_value_from_row(row, ["产品系列", "family"])},
                {"en": "CERTIFICATION", "cn": "认证标准", "value": _find_value_from_row(row, ["认证标准", "certification"])},
                {"en": "LUMEN", "cn": "流明输出", "value": _find_value_from_row(row, ["流明", "lumen"])},
                {"en": "EFFICIENCY", "cn": "光效", "value": _find_value_from_row(row, ["光效", "efficiency"])},
            ],
        },
        {
            "en": "Housing",
            "cn": "灯体参数",
            "items": [
                {"en": "HOUSING", "cn": "灯体", "value": _find_value_from_row(row, ["灯体", "housing"])},
                {"en": "COLOR", "cn": "颜色", "value": _find_value_from_row(row, ["颜色", "color"])},
                {"en": "SURFACE FINISH", "cn": "表面处理", "value": _find_value_from_row(row, ["表面处理", "surface finish"])},
                {"en": "IP", "cn": "防护等级", "value": _find_value_from_row(row, ["防护等级", " ip"])},
                {"en": "MOUNTING", "cn": "安装方式", "value": _find_value_from_row(row, ["安装方式", "mounting"])},
                {"en": "ACCESSORIES", "cn": "机械配件", "value": _find_value_from_row(row, ["机械配件", "accessories"])},
                {"en": "WEIGHT", "cn": "重量", "value": _find_value_from_row(row, ["重量", "weight"])},
            ],
        },
        {
            "en": "Optical",
            "cn": "光学参数",
            "items": [
                {"en": "TEMPERATURE", "cn": "工作温度", "value": _find_value_from_row(row, ["工作温度", "temperature"])},
                {"en": "SOURCE", "cn": "光源", "value": _find_value_from_row(row, ["光源", "source"])},
                {"en": "OPTICAL LENS", "cn": "光学类型", "value": _find_value_from_row(row, ["光学类型", "optical lens"])},
                {"en": "OPTIC ACCESSORIES", "cn": "光学配件", "value": _find_value_from_row(row, ["光学配件", "optic accessories"])},
                {"en": "BEAM", "cn": "光束角", "value": _find_value_from_row(row, ["光束角", "beam"])},
                {"en": "INTENSITY", "cn": "峰值光强", "value": _find_value_from_row(row, ["峰值光强", "intensity"])},
                {"en": "COLOR/CCT", "cn": "色温/色品", "value": _find_value_from_row(row, ["色彩/色温", "color/cct", "色温"])},
                {"en": "SDCM", "cn": "色容差", "value": _find_value_from_row(row, ["色容差", "sdcm"])},
                {"en": "RA", "cn": "显色指数", "value": _find_value_from_row(row, ["显色指数", " ra"])},
                {"en": "MAINTENANCE", "cn": "光通维持寿命", "value": _find_value_from_row(row, ["光通维持寿命", "maintenance"])},
            ],
        },
        {
            "en": "Electrical",
            "cn": "电气参数",
            "items": [
                {"en": "INPUT VOLTAGE", "cn": "输入电压", "value": _find_value_from_row(row, ["输入电压", "input voltage"])},
                {"en": "POWER", "cn": "额定功率", "value": _find_value_from_row(row, ["额定功率", "power"])},
                {"en": "CONTROL", "cn": "控制", "value": _find_value_from_row(row, ["控制", "control"])},
                {"en": "REQUIREMENT", "cn": "调光要求", "value": _find_value_from_row(row, ["调光要求", "requirement"])},
                {"en": "RATING", "cn": "电器防护等级", "value": _find_value_from_row(row, ["电器防护等级", "insulation", "rating"])},
                {"en": "CONNECTION", "cn": "灯具连接", "value": _find_value_from_row(row, ["灯具连接", "connection"])},
            ],
        },
    ]


@st.cache_data(show_spinner=False)
def load_image_mapping_from_assets(assets_dir: Path, image_subdir: str | None = None) -> dict[int, dict[str, object]]:
    if not assets_dir.exists():
        return {}

    # 如果指定了子目录，扫描子文件夹；否则扫描根目录（兼容旧结构）
    if image_subdir:
        product_dir = assets_dir / "产品图" / image_subdir
        dimension_dir = assets_dir / "尺寸图" / image_subdir
    else:
        product_dir = _resolve_candidate_dir(assets_dir, PRODUCT_IMAGE_DIR_CANDIDATES)
        dimension_dir = _resolve_candidate_dir(assets_dir, DIMENSION_IMAGE_DIR_CANDIDATES)

    supported_ext = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
    product_by_code: dict[int, Path] = {}
    dimension_by_code: dict[int, Path] = {}

    if product_dir and product_dir.exists():
        for file_path in product_dir.iterdir():
            if not file_path.is_file() or file_path.suffix.lower() not in supported_ext:
                continue
            for code in _extract_codes_from_filename(file_path):
                if code not in product_by_code:
                    product_by_code[code] = file_path

    if dimension_dir and dimension_dir.exists():
        for file_path in dimension_dir.iterdir():
            if not file_path.is_file() or file_path.suffix.lower() not in supported_ext:
                continue
            for code in _extract_codes_from_filename(file_path):
                if code not in dimension_by_code:
                    dimension_by_code[code] = file_path

    all_codes = sorted(set(product_by_code.keys()) | set(dimension_by_code.keys()))
    code_to_images: dict[int, dict[str, object]] = {}
    for code in all_codes:
        product_src = product_by_code.get(code)
        dimension_src = dimension_by_code.get(code)
        product_jpg = _normalize_image_to_jpg(product_src, "products") if product_src else None
        dimension_jpg = _normalize_image_to_jpg(dimension_src, "dimensions") if dimension_src else None
        code_to_images[code] = {
            "product_image_path": str(product_jpg.resolve()) if product_jpg else None,
            "dimension_image_path": str(dimension_jpg.resolve()) if dimension_jpg else None,
            "product_source_path": str(product_src.resolve()) if product_src else None,
            "dimension_source_path": str(dimension_src.resolve()) if dimension_src else None,
        }
    return code_to_images


@st.cache_data(show_spinner=False)
def load_data(file_path: Path) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=SHEET_NAME, header=None)
    en_headers = raw.iloc[HEADER_ROW_INDEX]
    cn_headers = raw.iloc[CN_HEADER_ROW_INDEX]

    columns: list[str] = []
    for idx, (en, cn) in enumerate(zip(en_headers, cn_headers)):
        en_clean = _clean_text(en)
        cn_clean = _clean_text(cn)
        if not en_clean and not cn_clean:
            columns.append(f"COL_{idx}")
            continue
        if en_clean and cn_clean:
            columns.append(f"{cn_clean} ({en_clean})")
        else:
            columns.append(en_clean or cn_clean)

    df = raw.iloc[DATA_START_ROW_INDEX:].copy()
    df.columns = columns

    # 删除全空列和全空行
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    # 清理文本中的换行和首尾空格
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(_clean_text).replace("", pd.NA)

    # 把可转为数字的字段转为数值，便于范围筛选
    for col in df.columns:
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().sum() > 0:
            non_null_original = df[col].notna().sum()
            if non_null_original > 0 and converted.notna().sum() >= non_null_original * 0.7:
                df[col] = converted

    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_all_data(excel_dir: Path) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for cat in CATEGORIES:
        file_path = excel_dir / cat["file"]  # type: ignore[arg-type]
        if not file_path.exists():
            continue
        df = load_data(file_path)
        df["产品分类 (CATEGORY)"] = cat["name"]
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    # 以第一个文件的列顺序为基准，缺少的列补 NA
    base_cols = frames[0].columns.tolist()
    aligned: list[pd.DataFrame] = []
    for f in frames:
        for c in base_cols:
            if c not in f.columns:
                f[c] = pd.NA
        aligned.append(f[base_cols])
    return pd.concat(aligned, ignore_index=True)


def main() -> None:
    st.set_page_config(page_title="灯具自动选型工具", layout="wide")
    st.title("灯具自动选型工具")
    st.caption("全品类产品库检索与规格书自动生成")

    try:
        df = load_all_data(EXCEL_DIR)
    except Exception as exc:
        st.exception(exc)
        return

    if df.empty:
        st.error("未加载到任何产品数据，请检查 excel/ 目录。")
        return

    # 合并所有分类的图片映射
    image_mapping: dict[int, dict[str, object]] = {}
    for cat in CATEGORIES:
        assets_dir: Path | None = cat["assets"]  # type: ignore[assignment]
        image_subdir: str | None = cat.get("image_subdir")  # type: ignore[assignment]
        if assets_dir:
            image_mapping.update(load_image_mapping_from_assets(assets_dir, image_subdir))

    st.success(f"已加载全品类 {len(df)} 条产品数据（{len(CATEGORIES)} 个分类），字段数 {len(df.columns)}")
    if image_mapping:
        st.caption(f"图片关联已启用：按 Code 可匹配 {len(image_mapping)} 条图片映射")

    type_col = _find_column(df.columns.tolist(), ["灯具类型", "type"])
    family_col = _find_column(df.columns.tolist(), ["产品系列", "family"])
    power_col = _find_column(df.columns.tolist(), ["额定功率", "power"])
    efficacy_col = _find_column(df.columns.tolist(), ["光效", "efficiency"])
    beam_col = _find_column(df.columns.tolist(), ["光束角", "beam"])
    cct_col = _find_column(df.columns.tolist(), ["色温", "color/cct", "色彩/色温"])
    nc12_col = _find_column(df.columns.tolist(), ["12nc"])

    required_cols = {
        "灯具类型": type_col,
        "产品系列": family_col,
        "额定功率": power_col,
        "光效": efficacy_col,
        "光束角": beam_col,
        "色温": cct_col,
        "12NC": nc12_col,
    }
    missing = [name for name, col in required_cols.items() if col is None]
    if missing:
        st.error(f"Excel 缺少必要字段: {', '.join(missing)}")
        return

    with st.sidebar:
        st.header("筛选条件")
        keep_null_numeric = st.checkbox("包含功率/光效为空的产品", value=True)

        filtered_df = df.copy()

        # 0) 产品分类
        cat_col = "产品分类 (CATEGORY)"
        cat_options = sorted(filtered_df[cat_col].dropna().astype(str).unique().tolist())
        selected_cats = st.multiselect("产品分类", options=cat_options)
        if selected_cats:
            filtered_df = filtered_df[filtered_df[cat_col].astype(str).isin(selected_cats)]

        # 1) 灯具类型
        type_options = sorted(filtered_df[type_col].dropna().astype(str).unique().tolist())
        selected_types = st.multiselect("灯具类型", options=type_options)
        if selected_types:
            filtered_df = filtered_df[filtered_df[type_col].astype(str).isin(selected_types)]

        # 2) 产品系列
        family_options = sorted(filtered_df[family_col].dropna().astype(str).unique().tolist())
        selected_families = st.multiselect("产品系列", options=family_options)
        if selected_families:
            filtered_df = filtered_df[filtered_df[family_col].astype(str).isin(selected_families)]

        # 3) 额定功率（下拉多选）
        power_numeric = pd.to_numeric(filtered_df[power_col], errors="coerce")
        power_options = sorted(power_numeric.dropna().unique().tolist())
        selected_powers = st.multiselect(
            "额定功率",
            options=power_options,
            format_func=lambda x: str(int(x)) if float(x).is_integer() else f"{float(x):g}",
        )
        if selected_powers:
            power_mask = power_numeric.isin(selected_powers)
            if keep_null_numeric:
                power_mask = power_mask | power_numeric.isna()
            filtered_df = filtered_df[power_mask]

        # 4) 色温
        cct_options = sorted(filtered_df[cct_col].dropna().astype(str).unique().tolist())
        selected_cct = st.multiselect("色温", options=cct_options)
        if selected_cct:
            filtered_df = filtered_df[filtered_df[cct_col].astype(str).isin(selected_cct)]

        # 5) 光效（范围）
        efficacy_series = pd.to_numeric(filtered_df[efficacy_col], errors="coerce").dropna()
        if not efficacy_series.empty and float(efficacy_series.min()) != float(efficacy_series.max()):
            eff_min, eff_max = float(efficacy_series.min()), float(efficacy_series.max())
            selected_eff_range = st.slider(
                "光效 范围",
                min_value=eff_min,
                max_value=eff_max,
                value=(eff_min, eff_max),
            )
            efficacy_numeric = pd.to_numeric(filtered_df[efficacy_col], errors="coerce")
            eff_mask = efficacy_numeric.between(selected_eff_range[0], selected_eff_range[1], inclusive="both")
            if keep_null_numeric:
                eff_mask = eff_mask | efficacy_numeric.isna()
            filtered_df = filtered_df[eff_mask]

        # 6) 光束角
        beam_options = sorted(filtered_df[beam_col].dropna().astype(str).unique().tolist())
        selected_beams = st.multiselect("光束角", options=beam_options)
        if selected_beams:
            filtered_df = filtered_df[filtered_df[beam_col].astype(str).isin(selected_beams)]

    col1, col2 = st.columns([1, 1])
    with col1:
        st.metric("筛选后条数", len(filtered_df))
    with col2:
        st.metric("筛选命中率", f"{(len(filtered_df) / len(df) * 100):.1f}%" if len(df) else "0%")

    if filtered_df.empty:
        st.warning("当前筛选条件没有匹配结果，请调整筛选项。")
        return

    st.subheader("候选结果（完整字段）")
    st.dataframe(filtered_df, use_container_width=True, height=500)

    st.subheader("规格书 PDF 生成")
    code_col = _find_column(filtered_df.columns.tolist(), ["编号", "code"])
    model_col = _find_column(filtered_df.columns.tolist(), ["型号", "model"])
    if code_col is None:
        st.warning("未找到 Code 列，暂无法生成规格书 PDF。")
    else:
        select_options: list[tuple[int, str]] = []
        for idx, row in filtered_df.iterrows():
            code_val = _to_display_text(row[code_col])
            model_val = _to_display_text(row[model_col]) if model_col else "/"
            select_options.append((idx, f"Code {code_val} | {model_val}"))

        selected_idx = st.selectbox(
            "选择需要生成规格书的产品",
            options=[item[0] for item in select_options],
            format_func=lambda x: next((txt for idx, txt in select_options if idx == x), str(x)),
        )
        selected_row = filtered_df.loc[selected_idx]

        project_name = st.text_input("项目名称", value="XXX项目")
        remarks_text = st.text_area("备注", value="/", height=80)

        if st.button("生成规格书 PDF", type="primary"):
            try:
                try:
                    from weasyprint import HTML
                except Exception as import_exc:
                    st.error("WeasyPrint 依赖未就绪，当前环境缺少系统库（如 GTK/Pango/Cairo）。")
                    st.exception(import_exc)
                    st.info("请先安装 WeasyPrint 的 Windows 系统依赖后再生成 PDF。")
                    return

                if not TEMPLATE_DIR.joinpath(SPEC_TEMPLATE_FILE).exists():
                    st.error(f"未找到模板文件: {TEMPLATE_DIR / SPEC_TEMPLATE_FILE}")
                else:
                    code_int = _to_code_int(selected_row[code_col])
                    image_info = image_mapping.get(code_int, {}) if code_int is not None else {}

                    product_image_path = image_info.get("product_image_path")
                    dimension_image_path = image_info.get("dimension_image_path")
                    distribution_image_path = None

                    env = Environment(loader=FileSystemLoader(str(TEMPLATE_DIR.resolve())))
                    template = env.get_template(SPEC_TEMPLATE_FILE)

                    html_content = template.render(
                        logo_path=LOGO_IMAGE_FILE.as_uri() if LOGO_IMAGE_FILE.exists() else None,
                        project_name=project_name,
                        product_code=_to_display_text(selected_row[code_col]),
                        version="1.0",
                        rev="1.0",
                        product_image_path=Path(product_image_path).as_uri() if product_image_path else None,
                        dimension_image_path=Path(dimension_image_path).as_uri() if dimension_image_path else None,
                        distribution_image_path=Path(distribution_image_path).as_uri() if distribution_image_path else None,
                        remarks=remarks_text,
                        sections=_build_spec_sections(selected_row),
                    )

                    pdf_bytes = HTML(string=html_content, base_url=str(TEMPLATE_DIR.resolve())).write_pdf()
                    GENERATED_PDF_DIR.mkdir(parents=True, exist_ok=True)
                    pdf_path = GENERATED_PDF_DIR / f"spec_sheet_code_{_to_display_text(selected_row[code_col]).replace('/', '_')}.pdf"
                    pdf_path.write_bytes(pdf_bytes)

                    st.success(f"PDF 生成成功：{pdf_path}")
                    st.download_button(
                        label="下载规格书 PDF",
                        data=pdf_bytes,
                        file_name=pdf_path.name,
                        mime="application/pdf",
                    )
            except Exception as exc:
                st.exception(exc)

        with st.expander("图片匹配调试信息", expanded=False):
            code_int = _to_code_int(selected_row[code_col])
            image_info = image_mapping.get(code_int, {}) if code_int is not None else {}
            st.write(f"当前 Code: {_to_display_text(selected_row[code_col])}")
            st.write(f"产品图路径: {image_info.get('product_image_path', '未命中')}")
            st.write(f"尺寸图路径: {image_info.get('dimension_image_path', '未命中')}")
            st.write(f"产品图源文件: {image_info.get('product_source_path', '未命中')}")
            st.write(f"尺寸图源文件: {image_info.get('dimension_source_path', '未命中')}")

    csv_data = filtered_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="下载当前筛选结果（完整字段）CSV",
        data=csv_data,
        file_name="lighting_filtered_full_result.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()