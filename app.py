import streamlit as st
from openpyxl.styles import Font, Border, Side
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile
from copy import copy

st.set_page_config(page_title="BOM Comparison Tracker")

st.title("BOM Comparison Tracker")

# =====================================
# COLORS
# =====================================

LIGHT_RED_FILL = PatternFill(
    start_color="F87474",
    end_color="F87474",
    fill_type="solid"
)

YELLOW_FILL = PatternFill(
    start_color="FFFF00",
    end_color="FFFF00",
    fill_type="solid"
)

ORANGE_FILL = PatternFill(
    start_color="FFA500",
    end_color="FFA500",
    fill_type="solid"
)

# =====================================
# FILE UPLOAD
# =====================================

master_file = st.file_uploader(
    "Upload Master BOM",
    type=["xlsx"]
)

bom_file = st.file_uploader(
    "Upload CPIF / PPIF Workbook",
    type=["xlsx"]
)

# =====================================
# PROCESS
# =====================================

if master_file and bom_file:

    @st.cache_resource
    def load_excel(file):
        return load_workbook(file)

    master_wb = load_excel(master_file)
    bom_wb = load_excel(bom_file)

    st.subheader("Sheet Selection")

    master_sheet = st.selectbox(
        "Master Sheet",
        master_wb.sheetnames
    )
    cpif_sheet = st.selectbox(
        "CPIF Sheet",
        bom_wb.sheetnames
    )
    ppif_sheet = st.selectbox(
        "PPIF Sheet",
        bom_wb.sheetnames
    )

    master_ws = master_wb[master_sheet]
    cpif_ws = bom_wb[cpif_sheet]
    ppif_ws = bom_wb[ppif_sheet]

    # =====================================
    # READ MASTER HEADERS (ROW 1)
    # =====================================

    headers = []

    for cell in master_ws[2]:
        if cell.value is not None:
            headers.append(str(cell.value).strip())

    st.subheader("Column Selection")

    key_column = st.selectbox(
        "Key Column",
        headers
    )

    compare_columns = st.multiselect(
        "Columns To Compare",
        headers
    )

    remark_columns = st.multiselect(
    "Columns To Show In Remarks",
    compare_columns,
    default=compare_columns
)

    cpif_filter = st.selectbox(
        "CPIF Buy Level Filter",
        ["All", "Y", "N"]
    )

    ppif_filter = st.selectbox(
        "PPIF Buy Level Filter",
        ["All", "Y", "N"]
    )

    # =====================================
    # COMPARE BUTTON
    # =====================================

    if st.button("Compare BOM"):

        master_headers = {}
        cpif_headers = {}
        ppif_headers = {}

        # Master headers row 1
        for cell in master_ws[2]:
            if cell.value is not None:
                master_headers[str(cell.value).strip()] = cell.column

        # CPIF headers row 1
        for cell in cpif_ws[1]:
                cpif_headers[
                    str(cell.value).strip()
                ] = cell.column

        # PPIF headers row 1
        for cell in ppif_ws[1]:
                ppif_headers[
                    str(cell.value).strip()
                ] = cell.column

        # =====================================
        # CREATE LOOKUPS
        # =====================================

        cpif_lookup = {}

        for r in range(2, cpif_ws.max_row + 1):

            if "ITEM NUMBER" not in cpif_headers:
                continue

            item_no = cpif_ws.cell(
                r,
                cpif_headers["ITEM NUMBER"]
            ).value

            if item_no is not None:
                cpif_lookup[
                    str(item_no).strip()
                ] = r

        ppif_lookup = {}

        for r in range(2, ppif_ws.max_row + 1):

            if "ITEM NUMBER" not in ppif_headers:
                continue

            item_no = ppif_ws.cell(
                r,
                ppif_headers["ITEM NUMBER"]
            ).value

            if item_no is not None:
                ppif_lookup[
                    str(item_no).strip()
                ] = r

        cpif_missing = 0
        ppif_missing = 0
        cpif_mismatch = 0
        ppif_mismatch = 0

        # =====================================
        # COMPARE MASTER BOM
        # =====================================

        last_header_col = 0

        for cell in master_ws[2]:
            if cell.value:
                last_header_col = cell.column

        # Insert a new column after Buy Level
        last_header_col=max(master_headers.values())
        remark_col = last_header_col + 1

        master_ws.cell(
            row=2,
            column=remark_col
        ).value = "Mismatch Remarks"

        

# Header format
        master_ws.cell(
            row=2,
            column=remark_col
        ).font = Font(
           bold=True,
           size=12
        )

        for r in range(2, master_ws.max_row + 1):

            source_cell = master_ws.cell(
                row=r,
                column=remark_col - 1   # SO 3 column
            )

            target_cell = master_ws.cell(
              row=r,
              column=remark_col
            )

    # Copy formatting from SO3 column
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.alignment = copy(source_cell.alignment)

    # Font
            if r == 2:
             target_cell.font = Font(
                 bold=True,
                 size=12
             )
            else:
                target_cell.font = Font(
                size=12
           )
   
        for master_row in range(4, master_ws.max_row + 1):
            remarks = []
            master_key = master_ws.cell(
                master_row,
                master_headers[key_column]
            ).value

            if master_key is None:
                continue

            master_key = str(master_key).strip()

            buy_level = ""

            if "Buy Level" in master_headers:

                buy_level = master_ws.cell(
                    master_row,
                    master_headers["Buy Level"]
                ).value

                buy_level = (
                    str(buy_level).strip().upper()
                    if buy_level
                    else ""
                )

            # ==============================
            # CPIF FILTER
            # ==============================

            cpif_check = True

            if cpif_filter == "Y":
                cpif_check = buy_level == "Y"

            elif cpif_filter == "N":
                cpif_check = buy_level == "N"

            # ==============================
            # PPIF FILTER
            # ==============================

            ppif_check = True

            if ppif_filter == "Y":
                ppif_check = buy_level == "Y"

            elif ppif_filter == "N":
                ppif_check = buy_level == "N"

            # ==============================
            # MISSING IN CPIF
            # ==============================

            if cpif_check:

               if master_key not in cpif_lookup:

                   master_ws.cell(
                       master_row,
                       master_headers[key_column]
                   ).fill = LIGHT_RED_FILL

                   if "Missing in CPIF" not in remarks:
                       remarks.append("Missing in CPIF")

                       cpif_missing += 1

            # ==============================
            # MISSING IN PPIF
            # ==============================

            if ppif_check:

               if master_key not in ppif_lookup:

                  master_ws.cell(
                      master_row,
                      master_headers[key_column]
                  ).fill = LIGHT_RED_FILL

                  if "Missing in PPIF" not in remarks:
                      remarks.append("Missing in PPIF")

                  ppif_missing += 1

            # ==============================
            # CPIF MISMATCH
            # ==============================

            if cpif_check and master_key in cpif_lookup:

                cpif_row = cpif_lookup[master_key]

                for col in compare_columns:

                    if (
                        col in master_headers
                        and col in cpif_headers
                    ):

                        master_val = master_ws.cell(
                            master_row,
                            master_headers[col]
                        ).value

                        cpif_val = cpif_ws.cell(
                            cpif_row,
                            cpif_headers[col]
                        ).value

                        master_val = str(master_val).strip() if master_val else ""
                        cpif_val = str(cpif_val).strip() if cpif_val else ""

                        if master_val != cpif_val:

                            master_ws.cell(
                                master_row,
                                master_headers[col]
                            ).fill = YELLOW_FILL

                            if f" CPIF {col} Mismatch" not in remarks:
                                 remarks.append(f" CPIF {col} Mismatch")

                            cpif_mismatch += 1

            # ==============================
            # PPIF MISMATCH
            # ==============================

            if ppif_check and master_key in ppif_lookup:

                ppif_row = ppif_lookup[master_key]

                for col in compare_columns:

                    if (
                        col in master_headers
                        and col in ppif_headers
                    ):

                        master_val = master_ws.cell(
                            master_row,
                            master_headers[col]
                        ).value

                        ppif_val = ppif_ws.cell(
                            ppif_row,
                            ppif_headers[col]
                        ).value

                        master_val = str(master_val).strip() if master_val else ""
                        ppif_val = str(ppif_val).strip() if ppif_val else ""

                        if master_val != ppif_val:

                            master_ws.cell(
                                master_row,
                                master_headers[col]
                            ).fill = ORANGE_FILL

                            if f" PPIF {col} Mismatch" not in remarks:
                                 remarks.append(f" PPIF {col} Mismatch")
                                

                            ppif_mismatch += 1

            # CREATE REQUIRED REMARK HEADERS
            # WRITE REMARKS
            master_ws.cell(
                row=master_row,
                column=remark_col
            ).value = ", ".join(remarks)

        # End of master_row loop
        # =====================================
        # SAVE RESULT
        # =====================================

        temp_file = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )
        output_file = temp_file.name
        master_wb.save(output_file)
        
        st.success("Comparison Completed")
        st.write(f"CPIF Missing : {cpif_missing}")
        st.write(f"PPIF Missing : {ppif_missing}")
        st.write(f"CPIF Mismatch : {cpif_mismatch}")
        st.write(f"PPIF Mismatch : {ppif_mismatch}")

        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Highlighted Master BOM",
                data=f,
                file_name="Highlighted_Master_BOM.xlsx"
            )
