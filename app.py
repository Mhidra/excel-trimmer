import streamlit as st
import openpyxl
import xlrd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
import io
import os

st.set_page_config(
    page_title="Excel Trimmer",
    page_icon="✂️",
    layout="centered"
)

st.title("✂️ Excel Trimmer")
st.markdown("Upload a large `.xlsx` or `.xls` file → get a trimmed version under 100 MB, ready for Google Sheets. **Nothing inside is changed** — only empty overflow rows/columns are removed.")

st.divider()

uploaded_file = st.file_uploader(
    "Drop your file here",
    type=["xlsx", "xls"],
    help="Supports .xlsx and .xls files up to 400 MB"
)

if uploaded_file:
    file_size_mb = round(uploaded_file.size / 1024 / 1024, 2)
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    st.info(f"📄 **{uploaded_file.name}** — {file_size_mb} MB")

    if st.button("🚀 Trim File", type="primary"):
        with st.spinner("Processing... this may take a minute for large files."):
            try:
                file_bytes = uploaded_file.read()

                if ext == "xlsx":
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        max_row = 0
                        max_col = 0
                        for row in ws.iter_rows():
                            for cell in row:
                                if cell.value is not None:
                                    if cell.row > max_row:
                                        max_row = cell.row
                                    if cell.column > max_col:
                                        max_col = cell.column
                        if max_row == 0:
                            continue
                        if ws.max_row > max_row:
                            ws.delete_rows(max_row + 1, ws.max_row - max_row)
                        if ws.max_column > max_col:
                            ws.delete_cols(max_col + 1, ws.max_column - max_col)

                else:  # .xls — full conversion preserving formatting
                    def xlrd_color(wb_xls, idx):
                        if idx in (0x7FFF, 64, 65, -1, None):
                            return None
                        try:
                            r, g, b = wb_xls.colour_map.get(idx, (None, None, None))
                            return f"{r:02X}{g:02X}{b:02X}" if r is not None else None
                        except:
                            return None

                    def border_style(n):
                        return {1:"thin",2:"medium",3:"dashed",4:"dotted",5:"thick",6:"double",
                                7:"hair",8:"mediumDashed",9:"dashDot",10:"mediumDashDot",
                                11:"dashDotDot",12:"mediumDashDotDot",13:"slantDashDot"}.get(n)

                    def make_side(wb_xls, bs, ci):
                        s = border_style(bs)
                        if not s:
                            return Side()
                        c = xlrd_color(wb_xls, ci)
                        return Side(border_style=s, color=c) if c else Side(border_style=s)

                    wb_xls = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)

                    for sheet_xls in wb_xls.sheets():
                        ws = wb.create_sheet(title=sheet_xls.name)
                        for ci in range(sheet_xls.ncols):
                            info = sheet_xls.colinfo_map.get(ci)
                            if info and info.width > 0:
                                ws.column_dimensions[get_column_letter(ci + 1)].width = info.width / 256.0
                        for ri in range(sheet_xls.nrows):
                            info = sheet_xls.rowinfo_map.get(ri)
                            if info and info.height:
                                ws.row_dimensions[ri + 1].height = info.height / 20.0
                        for rlo, rhi, clo, chi in sheet_xls.merged_cells:
                            if rhi - rlo > 1 or chi - clo > 1:
                                ws.merge_cells(start_row=rlo+1, start_column=clo+1, end_row=rhi, end_column=chi)
                        for ri in range(sheet_xls.nrows):
                            for ci in range(sheet_xls.ncols):
                                cell_xls = sheet_xls.cell(ri, ci)
                                cell_out = ws.cell(row=ri + 1, column=ci + 1)
                                if isinstance(cell_out, MergedCell):
                                    continue
                                t = cell_xls.ctype
                                if t == 1: cell_out.value = cell_xls.value
                                elif t == 2: cell_out.value = cell_xls.value
                                elif t == 3:
                                    try: cell_out.value = xlrd.xldate_as_datetime(cell_xls.value, wb_xls.datemode)
                                    except: cell_out.value = cell_xls.value
                                elif t == 4: cell_out.value = bool(cell_xls.value)
                                elif t not in (0, 6): cell_out.value = cell_xls.value
                                fmt_idx = sheet_xls.cell_xf_index(ri, ci)
                                xf = wb_xls.xf_list[fmt_idx]
                                font = wb_xls.font_list[xf.font_index]
                                fc = xlrd_color(wb_xls, font.colour_index)
                                cell_out.font = Font(
                                    name=font.name, size=font.height/20.0,
                                    bold=bool(font.bold), italic=bool(font.italic),
                                    underline="single" if font.underline_type else None,
                                    color=fc if fc else "000000")
                                bg = xf.background
                                if bg:
                                    fgc = xlrd_color(wb_xls, bg.pattern_colour_index)
                                    pt = getattr(bg, "pattern_type", getattr(bg, "fill_pattern", 0))
                                    if fgc and pt > 0:
                                        cell_out.fill = PatternFill(fill_type="solid", fgColor=fgc)
                                al = xf.alignment
                                cell_out.alignment = Alignment(
                                    horizontal={0:"general",1:"left",2:"center",3:"right",4:"fill",5:"justify",6:"centerContinuous",7:"distributed"}.get(al.hor_align,"general"),
                                    vertical={0:"top",1:"center",2:"bottom",3:"justify",4:"distributed"}.get(al.vert_align,"bottom"),
                                    wrap_text=bool(al.text_wrapped))
                                brd = xf.border
                                cell_out.border = Border(
                                    left=make_side(wb_xls, getattr(brd,"left_line_type",0), getattr(brd,"left_colour_index",0)),
                                    right=make_side(wb_xls, getattr(brd,"right_line_type",0), getattr(brd,"right_colour_index",0)),
                                    top=make_side(wb_xls, getattr(brd,"top_line_type",0), getattr(brd,"top_colour_index",0)),
                                    bottom=make_side(wb_xls, getattr(brd,"bottom_line_type",0), getattr(brd,"bottom_colour_index",0)))
                                fmt = wb_xls.format_map.get(xf.format_key)
                                if fmt:
                                    cell_out.number_format = fmt.format_str

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                result_bytes = output.getvalue()
                result_mb = round(len(result_bytes) / 1024 / 1024, 2)

                base_name = uploaded_file.name.rsplit(".", 1)[0]
                output_name = f"{base_name}_FINAL.xlsx"

                st.success(f"✅ Done! {file_size_mb} MB → **{result_mb} MB**")

                st.download_button(
                    label="⬇️ Download Trimmed File",
                    data=result_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.info("Make sure the file is a valid .xlsx or .xls file and try again.")

st.divider()
st.caption("Works with files up to ~400 MB. All processing happens on the server — your file is never stored.")
