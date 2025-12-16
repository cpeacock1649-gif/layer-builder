import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import itertools
import textwrap
import math
from database import InsuranceDB
from excel_export import export_program_to_excel
from word_export import export_program_to_word
from pdf_parser import parse_insurance_pdf, merge_parsed_documents
from excel_parser import parse_excel_program, merge_excel_programs


# Page config
st.set_page_config(
    page_title="Insurance Layer Builder",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Initialize database
if "db" not in st.session_state:
    st.session_state.db = InsuranceDB()

db = st.session_state.db

# Custom CSS
st.markdown(
    """
<style>
    /* Set base font for the entire application */
    html, body, [class*="st-"], .stMarkdown, .stExpander, div.stButton button, 
    .stSelectbox, .stNumberInput, .stTextInput, .stCheckbox, .stRadio,
    div.streamlit-expanderHeader, .streamlit-expanderContent {
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Ensure consistent font in expander headers */
    .streamlit-expanderHeader {
        font-weight: 400 !important;
        font-size: 1rem !important;
        letter-spacing: normal !important;
    }
    
    /* Fix for layer titles in expanders */
    .streamlit-expanderHeader p {
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
        font-weight: 400 !important;
        font-size: 1rem !important;
    }
    
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    
    .carrier-group {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin-bottom: 1rem;
    }
    
    /* Ensure consistent styling for buttons */
    div.stButton > button {
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
        font-weight: 400;
    }
    
    /* Ensure consistent styling for metrics */
    .css-1wivap2 {
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Fix for overlapping text in dashboard */
    [data-testid="stMetricLabel"] > div:first-child {
        overflow: hidden !important;
    }
    
    /* Hide only SVG icons in expander headers */
    details[data-testid="stExpander"] summary svg {
        display: none !important;
    }
    
    /* Hide any icon wrapper that contains broken text - be very specific */
    [data-testid="stExpanderToggleIcon"] {
        font-size: 0 !important;
        width: 1rem !important;
        overflow: hidden !important;
    }
    
    /* Replace with clean Unicode arrows */
    [data-testid="stExpanderToggleIcon"]::before {
        content: "▶";
        font-size: 0.75rem !important;
        display: inline-block !important;
    }
    
    details[data-testid="stExpander"][open] [data-testid="stExpanderToggleIcon"]::before {
        content: "▼";
    }
</style>
""",
    unsafe_allow_html=True,
)


def format_layer_title(layer, idx):
    """Format layer title with proper attachment logic"""
    limit_val = layer.get("limit", 0)
    attach_val = layer.get("attachment", 0)

    if layer.get("is_primary"):
        return f"Layer {idx + 1}: ${limit_val:,.0f} Primary"
    else:
        return f"Layer {idx + 1}: ${limit_val:,.0f} xs ${attach_val:,.0f}"


def styled_expander(title, expanded=False):
    """Custom styled expander that ensures consistent font and icon rendering"""
    # Add a space before the title to accommodate the arrow icon better
    styled_title = f"{title}"
    return st.expander(styled_title, expanded=expanded)


def styled_header(text, level=1):
    """Display a styled header with consistent font"""
    if level == 1:
        st.markdown(f"<h1 class='main-header'>{text}</h1>", unsafe_allow_html=True)
    else:
        st.markdown(
            f"<h{level} style='font-family: \"Helvetica Neue\", Arial, sans-serif;'>{text}</h{level}>",
            unsafe_allow_html=True,
        )


# Sidebar
with st.sidebar:
    st.title("🏢 Insurance Layer Builder")
    st.markdown("---")

    menu = st.radio(
        "Navigation",
        [
            "📊 Dashboard",
            "🔨 Build Program",
            "📚 Carrier Library",
            "⚙️ Settings",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.caption("v1.0.0 | Standalone Edition")

# Dashboard View
if menu == "📊 Dashboard":
    st.markdown(
        "<h1 class='main-header'>📊 Account Dashboard</h1>", unsafe_allow_html=True
    )

    accounts = db.get_all_accounts()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Accounts", len(accounts))
    with col2:
        carriers = db.get_all_carriers()
        st.metric("Registered Carriers", len(carriers))
    with col3:
        st.metric("Active Programs", len([a for a in accounts if db.get_program(a[0])]))

    st.markdown("---")

    # New Account Section
    with styled_expander("➕ Create New Account", expanded=False):
        new_account_name = st.text_input("Account Name")
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("Create Account", type="primary", use_container_width=True):
                if new_account_name:
                    result = db.add_account(new_account_name)
                    if result:
                        st.success(f"✅ Account '{new_account_name}' created!")
                        st.rerun()
                    else:
                        st.error("❌ Account already exists!")
                else:
                    st.warning("⚠️ Please enter an account name")

    st.markdown("### 📋 All Accounts")

    if accounts:
        for account in accounts:
            account_id, name, created, modified = account
            program = db.get_program(account_id)

            with styled_expander(f"**{name}**", expanded=False):
                col1, col2, col3, col4 = st.columns([2, 1, 1, 1])

                with col1:
                    layer_count = len(program.get("layers", [])) if program else 0
                    st.write(f"**Layers:** {layer_count}")

                with col2:
                    if st.button(
                        "✏️ Edit", key=f"edit_{account_id}", use_container_width=True
                    ):
                        st.session_state.selected_account_id = account_id
                        st.session_state.current_menu = "🔨 Build Program"
                        st.rerun()

                with col3:
                    if st.button(
                        "📋 Clone",
                        key=f"clone_{account_id}",
                        use_container_width=True,
                    ):
                        new_name = f"{name} (Copy)"
                        db.clone_account(account_id, new_name)
                        st.success(f"Cloned to '{new_name}'")
                        st.rerun()

                with col4:
                    if st.button(
                        "🗑️ Delete",
                        key=f"del_{account_id}",
                        use_container_width=True,
                    ):
                        db.delete_account(account_id)
                        st.success("Account deleted")
                        st.rerun()
    else:
        st.info("📝 No accounts yet. Create your first account above!")

# Build Program View
elif menu == "🔨 Build Program":
    st.markdown(
        "<h1 class='main-header'>🔨 Build Insurance Program</h1>",
        unsafe_allow_html=True,
    )

    accounts = db.get_all_accounts()

    if not accounts:
        st.warning("⚠️ No accounts available. Create one in the Dashboard first.")
    else:
        # Account selection
        col1, col2 = st.columns([3, 1])

        with col1:
            if "selected_account_id" not in st.session_state:
                st.session_state.selected_account_id = accounts[0][0]

            account_options = {acc[0]: acc[1] for acc in accounts}
            selected_account_id = st.selectbox(
                "Select Account",
                options=list(account_options.keys()),
                format_func=lambda x: account_options[x],
                index=(
                    list(account_options.keys()).index(
                        st.session_state.selected_account_id
                    )
                    if st.session_state.selected_account_id in account_options
                    else 0
                ),
            )
            st.session_state.selected_account_id = selected_account_id

        with col2:
            build_mode = st.toggle("🔧 Build Mode", value=True)

        # Load program
        program = db.get_program(selected_account_id)

        if program:
            if build_mode:
                st.info(
                    "🔧 **Build Mode Active** - Add, edit, or remove layers and carriers"
                )

                if "edited_program" not in st.session_state:
                    import copy

                    st.session_state.edited_program = copy.deepcopy(program)

                edited_program = st.session_state.edited_program

                # PDF Import Section
                st.markdown("---")
                with st.expander("📥 Import from PDF Documents", expanded=False):
                    st.markdown(
                        """
                    Upload insurance quotes, binders, or policies (up to 25 PDFs) to automatically extract:
                    - Policy limits and attachment points
                    - Carrier information and shares
                    - Premium amounts
                    - Policy numbers
                    - Carrier-specific limit allocations
                    """
                    )

                    uploaded_files = st.file_uploader(
                        "Upload PDF documents",
                        type=["pdf"],
                        accept_multiple_files=True,
                        key="pdf_uploader",
                        help="Upload up to 25 insurance quote or binder PDFs",
                    )

                    if uploaded_files:
                        if len(uploaded_files) > 25:
                            st.warning(
                                "⚠️ Maximum 25 files allowed. Only the first 25 will be processed."
                            )
                            uploaded_files = uploaded_files[:25]

                        col_parse, col_cancel = st.columns([1, 1])

                        with col_parse:
                            if st.button(
                                "🔍 Parse PDFs",
                                type="primary",
                                use_container_width=True,
                            ):
                                # Parse all uploaded PDFs
                                with st.spinner(
                                    f"Parsing {len(uploaded_files)} document(s)..."
                                ):
                                    parsed_results = []

                                    for uploaded_file in uploaded_files:
                                        try:
                                            pdf_bytes = uploaded_file.read()
                                            result = parse_insurance_pdf(
                                                pdf_bytes, uploaded_file.name
                                            )
                                            parsed_results.append(result)
                                        except Exception as e:
                                            st.error(
                                                f"Error parsing {uploaded_file.name}: {str(e)}"
                                            )

                                    # Store parsed results in session state
                                    st.session_state.parsed_pdfs = parsed_results

                        with col_cancel:
                            if st.button("❌ Clear Upload", use_container_width=True):
                                if "parsed_pdfs" in st.session_state:
                                    del st.session_state.parsed_pdfs
                                st.rerun()

                    # Display parsed results and allow user to review/import
                    if (
                        "parsed_pdfs" in st.session_state
                        and st.session_state.parsed_pdfs
                    ):
                        st.markdown("---")
                        st.markdown("#### 📊 Parsed Results")

                        parsed_docs = st.session_state.parsed_pdfs
                        success_count = len(
                            [d for d in parsed_docs if d.get("success")]
                        )
                        failed_count = len(
                            [d for d in parsed_docs if not d.get("success")]
                        )

                        st.info(
                            f"✅ Successfully parsed: {success_count} | ❌ Failed: {failed_count}"
                        )

                        # Show summary of each document
                        for i, doc in enumerate(parsed_docs):
                            if doc.get("success"):
                                with st.expander(
                                    f"📄 {doc.get('filename', f'Document {i+1}')} - {doc.get('document_type', 'Unknown')}",
                                    expanded=False,
                                ):
                                    col1, col2 = st.columns(2)

                                    with col1:
                                        st.markdown("**Limits Found:**")
                                        if doc.get("limits"):
                                            for limit in doc["limits"]:
                                                if limit["is_primary"]:
                                                    st.write(
                                                        f"- Primary: ${limit['limit']:,.0f}"
                                                    )
                                                else:
                                                    st.write(
                                                        f"- ${limit['limit']:,.0f} xs ${limit['attachment']:,.0f}"
                                                    )
                                        else:
                                            st.write("- None detected")

                                    with col2:
                                        st.markdown("**Carriers Found:**")
                                        if doc.get("carriers"):
                                            for carrier in doc["carriers"]:
                                                st.write(
                                                    f"- {carrier['carrier_name']}: {carrier['share']*100:.1f}%"
                                                )
                                        else:
                                            st.write("- None detected")

                                    # Show "part of" data if found
                                    if doc.get("part_of_data"):
                                        st.markdown("---")
                                        st.markdown("**Carrier-Specific Allocations:**")
                                        for part_of in doc["part_of_data"]:
                                            layer_desc = (
                                                f"Primary ${part_of['layer_limit']:,.0f}"
                                                if part_of["is_primary"]
                                                else f"${part_of['layer_limit']:,.0f} xs ${part_of['attachment']:,.0f}"
                                            )
                                            st.write(
                                                f"- **{part_of['carrier_name']}**: ${part_of['carrier_limit']:,.0f} ({part_of['share']*100:.2f}%) of {layer_desc}"
                                            )

                                    if doc.get("policy_number"):
                                        st.markdown(
                                            f"**Policy Number:** {doc['policy_number']}"
                                        )

                            else:
                                with st.expander(
                                    f"❌ {doc.get('filename', f'Document {i+1}')} - Failed",
                                    expanded=False,
                                ):
                                    st.error(
                                        f"Error: {doc.get('error', 'Unknown error')}"
                                    )

                        # Import options
                        st.markdown("---")
                        st.markdown("#### 🔨 Import Options")

                        col_import, col_append = st.columns(2)

                        with col_import:
                            if st.button(
                                "🔄 Replace Program with Parsed Data",
                                use_container_width=True,
                            ):
                                # Merge all parsed documents into program structure
                                merged_data = merge_parsed_documents(parsed_docs)

                                if merged_data["layers"]:
                                    edited_program["layers"] = merged_data["layers"]
                                    st.session_state.edited_program = edited_program
                                    st.success(
                                        f"✅ Program replaced! Imported {merged_data['documents_processed']} document(s) with {len(merged_data['layers'])} layer(s)"
                                    )
                                    # Clear parsed data
                                    del st.session_state.parsed_pdfs
                                    st.rerun()
                                else:
                                    st.warning(
                                        "⚠️ No valid layers found in parsed documents"
                                    )

                        with col_append:
                            if st.button(
                                "➕ Append to Existing Program",
                                use_container_width=True,
                            ):
                                # Merge parsed documents and append to existing layers
                                merged_data = merge_parsed_documents(parsed_docs)

                                if merged_data["layers"]:
                                    # Append new layers to existing ones
                                    edited_program["layers"].extend(
                                        merged_data["layers"]
                                    )
                                    st.session_state.edited_program = edited_program
                                    st.success(
                                        f"✅ Layers appended! Added {len(merged_data['layers'])} layer(s) from {merged_data['documents_processed']} document(s)"
                                    )
                                    # Clear parsed data
                                    del st.session_state.parsed_pdfs
                                    st.rerun()
                                else:
                                    st.warning(
                                        "⚠️ No valid layers found in parsed documents"
                                    )

                st.markdown("---")

                # Excel Import Section - expand by default if no layers exist
                excel_section_expanded = len(edited_program.get("layers", [])) == 0
                with st.expander(
                    "📊 Import from Excel Program Structure",
                    expanded=excel_section_expanded,
                ):
                    st.markdown(
                        """
                    **Upload broker program schedules** (like OHSU-style spreadsheets) to automatically extract:
                    - Insurance layers with limits and attachment points (e.g., "$75M ex $100M EQ")
                    - Carrier participation and line amounts from Participant rows
                    - Premium amounts, fees, and surplus lines taxes
                    - Shares calculated automatically from Line / Layer Limit
                    
                    **Supported formats:**
                    - Layer headers: "$75M ex $100M", "$500M ex $1BL", "Terrorism", etc.
                    - Column headers: Participant, Line, Premium, Fees, SL Tax, Total
                    """
                    )

                    uploaded_excel_files = st.file_uploader(
                        "Upload Excel files",
                        type=["xlsx", "xls"],
                        accept_multiple_files=True,
                        key="excel_uploader",
                        help="Upload up to 10 insurance program Excel files",
                    )

                    # Debug mode toggle
                    debug_parsing = st.checkbox(
                        "🔍 Enable Debug Mode",
                        value=False,
                        help="Show detailed parsing information to troubleshoot issues",
                    )

                    if uploaded_excel_files:
                        if len(uploaded_excel_files) > 10:
                            st.warning(
                                "⚠️ Maximum 10 files allowed. Only the first 10 will be processed."
                            )
                            uploaded_excel_files = uploaded_excel_files[:10]

                        col_parse_xl, col_cancel_xl = st.columns([1, 1])

                        with col_parse_xl:
                            if st.button(
                                "🔍 Parse Excel Files",
                                type="primary",
                                use_container_width=True,
                                key="parse_excel_btn",
                            ):
                                # Parse all uploaded Excel files
                                with st.spinner(
                                    f"Parsing {len(uploaded_excel_files)} Excel file(s)..."
                                ):
                                    parsed_excel_results = []
                                    debug_output = []

                                    for uploaded_file in uploaded_excel_files:
                                        try:
                                            excel_bytes = uploaded_file.read()
                                            result = parse_excel_program(
                                                excel_bytes,
                                                uploaded_file.name,
                                                debug=debug_parsing,
                                            )
                                            parsed_excel_results.append(result)

                                            # Collect debug info for display
                                            if debug_parsing and result.get("success"):
                                                debug_output.append(
                                                    f"**File: {result.get('filename', 'Unknown')}**"
                                                )
                                                for layer in result.get("layers", []):
                                                    layer_limit = layer.get("limit", 0)
                                                    layer_attach = layer.get(
                                                        "attachment", 0
                                                    )
                                                    carriers = layer.get("carriers", [])
                                                    total_share = sum(
                                                        c.get("share", 0)
                                                        for c in carriers
                                                    )

                                                    # Show validation status
                                                    status = (
                                                        "✅"
                                                        if abs(total_share - 1.0) < 0.02
                                                        else "⚠️"
                                                    )

                                                    # Count unique carrier names vs total entries
                                                    unique_names = len(
                                                        set(
                                                            c["carrier_name"]
                                                            for c in carriers
                                                        )
                                                    )
                                                    entries_text = (
                                                        f"{len(carriers)} entries"
                                                        if len(carriers) != unique_names
                                                        else f"{len(carriers)} carriers"
                                                    )

                                                    debug_output.append(
                                                        f"**Layer: ${layer_limit:,.0f} xs ${layer_attach:,.0f}** ({entries_text})"
                                                    )
                                                    for idx, c in enumerate(carriers):
                                                        line_calc = (
                                                            c.get("share", 0)
                                                            * layer_limit
                                                        )
                                                        debug_output.append(
                                                            f"  {idx+1}. {c['carrier_name']}: {c['share']*100:.2f}% (Line=${line_calc:,.0f})"
                                                        )
                                                    debug_output.append(
                                                        f"  {status} **Total: {total_share*100:.1f}%**"
                                                    )
                                                    debug_output.append("")

                                        except Exception as e:
                                            st.error(
                                                f"Error parsing {uploaded_file.name}: {str(e)}"
                                            )

                                    # Store parsed results in session state
                                    st.session_state.parsed_excel = parsed_excel_results

                                    # Show debug output if enabled
                                    if debug_parsing and debug_output:
                                        st.session_state.excel_debug_output = (
                                            debug_output
                                        )

                        with col_cancel_xl:
                            if st.button(
                                "❌ Clear Excel Upload",
                                use_container_width=True,
                                key="clear_excel_btn",
                            ):
                                if "parsed_excel" in st.session_state:
                                    del st.session_state.parsed_excel
                                if "excel_debug_output" in st.session_state:
                                    del st.session_state.excel_debug_output
                                st.rerun()

                    # Display parsed Excel results
                    if (
                        "parsed_excel" in st.session_state
                        and st.session_state.parsed_excel
                    ):
                        st.markdown("---")

                        # Show debug output if available
                        if (
                            "excel_debug_output" in st.session_state
                            and st.session_state.excel_debug_output
                        ):
                            with st.expander(
                                "🔍 Debug: Share Calculations", expanded=True
                            ):
                                for line in st.session_state.excel_debug_output:
                                    st.markdown(line)

                        st.markdown("#### 📊 Parsed Excel Results")

                        parsed_excel = st.session_state.parsed_excel
                        success_count = len(
                            [d for d in parsed_excel if d.get("success")]
                        )
                        failed_count = len(
                            [d for d in parsed_excel if not d.get("success")]
                        )

                        st.info(
                            f"✅ Successfully parsed: {success_count} | ❌ Failed: {failed_count}"
                        )

                        # Show summary of each Excel file
                        for i, excel_data in enumerate(parsed_excel):
                            if excel_data.get("success"):
                                with st.expander(
                                    f"📊 {excel_data.get('filename', f'Excel File {i+1}')}",
                                    expanded=False,
                                ):
                                    layers = excel_data.get("layers", [])
                                    st.markdown(f"**Layers Found:** {len(layers)}")

                                    for j, layer in enumerate(layers):
                                        layer_desc = (
                                            f"Primary: ${layer['limit']:,.0f}"
                                            if layer["is_primary"]
                                            else f"${layer['limit']:,.0f} xs ${layer['attachment']:,.0f}"
                                        )
                                        st.markdown(f"**Layer {j+1}:** {layer_desc}")
                                        st.write(
                                            f"  - Carriers: {len(layer.get('carriers', []))}"
                                        )

                                        if layer.get("carriers"):
                                            # Calculate and show total share for validation
                                            total_share = sum(
                                                c.get("share", 0)
                                                for c in layer["carriers"]
                                            )
                                            share_status = (
                                                "✅"
                                                if abs(total_share - 1.0) < 0.01
                                                else "⚠️"
                                            )
                                            st.write(
                                                f"  - Total Share: {share_status} {total_share*100:.1f}%"
                                            )

                                            for carrier in layer["carriers"][
                                                :5
                                            ]:  # Show first 5 carriers
                                                # Calculate line amount from share for verification
                                                line_amount = (
                                                    carrier.get("share", 0)
                                                    * layer["limit"]
                                                )
                                                premium_str = (
                                                    f"${carrier.get('premium', 0):,.0f}"
                                                    if carrier.get("premium", 0) > 0
                                                    else "N/A"
                                                )
                                                st.write(
                                                    f"    • {carrier['carrier_name']}: {carrier['share']*100:.2f}% (Line: ${line_amount:,.0f}) - Premium: {premium_str}"
                                                )

                                            if len(layer["carriers"]) > 5:
                                                st.write(
                                                    f"    ... and {len(layer['carriers']) - 5} more carriers"
                                                )

                            else:
                                with st.expander(
                                    f"❌ {excel_data.get('filename', f'Excel File {i+1}')} - Failed",
                                    expanded=False,
                                ):
                                    st.error(
                                        f"Error: {excel_data.get('error', 'Unknown error')}"
                                    )

                        # Import options
                        st.markdown("---")
                        st.markdown("#### 🔨 Excel Import Options")

                        col_import_xl, col_append_xl = st.columns(2)

                        with col_import_xl:
                            if st.button(
                                "🔄 Replace Program with Excel Data",
                                use_container_width=True,
                                key="replace_excel_btn",
                                type="primary",
                            ):
                                # Merge all parsed Excel files
                                merged_data = merge_excel_programs(parsed_excel)

                                if merged_data["layers"]:
                                    edited_program["layers"] = merged_data["layers"]
                                    st.session_state.edited_program = edited_program

                                    # Calculate total carriers
                                    total_carriers = sum(
                                        len(layer.get("carriers", []))
                                        for layer in merged_data["layers"]
                                    )

                                    st.success(
                                        f"✅ Program imported! {merged_data['documents_processed']} file(s) → "
                                        f"{len(merged_data['layers'])} layers with {total_carriers} carrier entries. "
                                        f"**Scroll down to see the Mud Map visualization! ↓**"
                                    )
                                    # Clear parsed data
                                    del st.session_state.parsed_excel
                                    if "excel_debug_output" in st.session_state:
                                        del st.session_state.excel_debug_output
                                    st.rerun()
                                else:
                                    st.warning(
                                        "⚠️ No valid layers found in Excel files. Check that layer headers like '$75M ex $100M' are present."
                                    )

                        with col_append_xl:
                            if st.button(
                                "➕ Append Excel Data to Program",
                                use_container_width=True,
                                key="append_excel_btn",
                            ):
                                # Merge Excel files and append to existing layers
                                merged_data = merge_excel_programs(parsed_excel)

                                if merged_data["layers"]:
                                    # Append new layers to existing ones
                                    edited_program["layers"].extend(
                                        merged_data["layers"]
                                    )
                                    st.session_state.edited_program = edited_program

                                    # Calculate total carriers
                                    total_carriers = sum(
                                        len(layer.get("carriers", []))
                                        for layer in merged_data["layers"]
                                    )

                                    st.success(
                                        f"✅ Layers appended! Added {len(merged_data['layers'])} layer(s) with {total_carriers} carriers. "
                                        f"**Scroll down to see the Mud Map! ↓**"
                                    )
                                    # Clear parsed data
                                    del st.session_state.parsed_excel
                                    if "excel_debug_output" in st.session_state:
                                        del st.session_state.excel_debug_output
                                    st.rerun()
                                else:
                                    st.warning(
                                        "⚠️ No valid layers found in Excel files. Check that layer headers like '$75M ex $100M' are present."
                                    )

                st.markdown("---")

                # Layer Management
                st.markdown("### 📚 Layer Management")

                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("➕ Add New Layer", use_container_width=True):
                        # Calculate next attachment point
                        if edited_program["layers"]:
                            sorted_layers = sorted(
                                edited_program["layers"],
                                key=lambda x: x.get("attachment", 0),
                            )
                            last_layer = sorted_layers[-1]
                            next_attachment = last_layer.get(
                                "attachment", 0
                            ) + last_layer.get("limit", 0)
                        else:
                            next_attachment = 0

                        new_layer = {
                            "limit": 1000000,
                            "attachment": next_attachment,
                            "is_primary": len(edited_program["layers"]) == 0,
                            "carriers": [],
                        }
                        edited_program["layers"].append(new_layer)
                        st.rerun()

                with col2:
                    if edited_program["layers"] and st.button(
                        "📋 Duplicate Last Layer", use_container_width=True
                    ):
                        import copy

                        sorted_layers = sorted(
                            edited_program["layers"],
                            key=lambda x: x.get("attachment", 0),
                        )
                        last_layer = sorted_layers[-1]
                        new_layer = copy.deepcopy(last_layer)
                        # Calculate next attachment
                        new_layer["attachment"] = last_layer.get(
                            "attachment", 0
                        ) + last_layer.get("limit", 0)
                        new_layer["is_primary"] = False
                        edited_program["layers"].append(new_layer)
                        st.rerun()

                with col3:
                    template = st.selectbox(
                        "Quick Template",
                        [
                            "None",
                            "Primary Only",
                            "Primary + 1 Excess",
                            "Primary + 2 Excess",
                        ],
                    )
                    if template != "None" and st.button(
                        "Apply Template", use_container_width=True
                    ):
                        edited_program["layers"] = []
                        if template == "Primary Only":
                            edited_program["layers"].append(
                                {
                                    "limit": 1000000,
                                    "attachment": 0,
                                    "is_primary": True,
                                    "carriers": [],
                                }
                            )
                        elif template == "Primary + 1 Excess":
                            edited_program["layers"].append(
                                {
                                    "limit": 1000000,
                                    "attachment": 0,
                                    "is_primary": True,
                                    "carriers": [],
                                }
                            )
                            edited_program["layers"].append(
                                {
                                    "limit": 5000000,
                                    "attachment": 1000000,
                                    "is_primary": False,
                                    "carriers": [],
                                }
                            )
                        elif template == "Primary + 2 Excess":
                            edited_program["layers"].append(
                                {
                                    "limit": 1000000,
                                    "attachment": 0,
                                    "is_primary": True,
                                    "carriers": [],
                                }
                            )
                            edited_program["layers"].append(
                                {
                                    "limit": 5000000,
                                    "attachment": 1000000,
                                    "is_primary": False,
                                    "carriers": [],
                                }
                            )
                            edited_program["layers"].append(
                                {
                                    "limit": 10000000,
                                    "attachment": 6000000,
                                    "is_primary": False,
                                    "carriers": [],
                                }
                            )
                        st.rerun()

                # Display layers
                layers_to_delete = []
                sorted_layers = sorted(
                    enumerate(edited_program["layers"]),
                    key=lambda x: x[1].get("attachment", 0),
                )

                for display_idx, (original_idx, layer) in enumerate(sorted_layers):
                    layer_title = format_layer_title(layer, display_idx)

                    # Ensure carriers list exists
                    if "carriers" not in layer:
                        layer["carriers"] = []

                    with styled_expander(layer_title, expanded=True):
                        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])

                        with col1:
                            new_limit = st.number_input(
                                "Limit ($)",
                                value=float(layer.get("limit", 0)),
                                key=f"limit_{original_idx}",
                                format="%.0f",
                                min_value=0.0,
                            )
                            layer["limit"] = new_limit

                        with col2:
                            new_attachment = st.number_input(
                                "Attachment ($)",
                                value=float(layer.get("attachment", 0)),
                                key=f"attach_{original_idx}",
                                format="%.0f",
                                disabled=layer.get("is_primary", False),
                                min_value=0.0,
                            )
                            if not layer.get("is_primary"):
                                layer["attachment"] = new_attachment

                        with col3:
                            is_primary = st.checkbox(
                                "Primary",
                                value=layer.get("is_primary", False),
                                key=f"primary_{original_idx}",
                            )
                            layer["is_primary"] = is_primary

                        with col4:
                            if st.button(
                                "🗑️ Delete Layer",
                                key=f"del_layer_{original_idx}",
                                use_container_width=True,
                            ):
                                layers_to_delete.append(original_idx)

                        if layer.get("is_primary"):
                            layer["attachment"] = 0

                        # Carrier Management
                        st.markdown("#### 🏢 Carriers")

                        if st.button(
                            f"➕ Add Carrier",
                            key=f"add_carrier_{original_idx}",
                            use_container_width=True,
                        ):
                            layer["carriers"].append(
                                {
                                    "carrier_name": "",
                                    "share": 0.0,  # Combined share/carrier_percent
                                    "premium": 0,
                                    "carrier_fee": 0.0,
                                    "surplus_fee": 0.0,
                                    "policy_number": "",
                                    "has_multiple_rbes": False,
                                    "rbes": [],
                                }
                            )
                            st.rerun()

                        carriers_to_delete = []

                        for cidx, carrier in enumerate(layer["carriers"]):
                            st.markdown(
                                f"<div class='carrier-group'>", unsafe_allow_html=True
                            )

                            # Carrier header with delete button
                            col_carrier, col_multi_rbe, col_del = st.columns([3, 1, 1])
                            with col_carrier:
                                carrier_name = carrier.get(
                                    "carrier_name", "New Carrier"
                                )
                                st.markdown(f"**🏢 {carrier_name}**")
                            with col_multi_rbe:
                                has_multiple = carrier.get("has_multiple_rbes", False)
                                if st.button(
                                    (
                                        "📋 Multiple RBE"
                                        if not has_multiple
                                        else "✅ Multiple RBE"
                                    ),
                                    key=f"multi_rbe_{original_idx}_{cidx}",
                                    use_container_width=True,
                                ):
                                    carrier["has_multiple_rbes"] = not has_multiple
                                    # If enabling multiple RBEs, initialize with current data
                                    if not has_multiple and not carrier.get("rbes"):
                                        carrier["rbes"] = [
                                            {
                                                "rbe": "",
                                                "share": 1.0,  # Default to 100% of carrier's share
                                                "premium": carrier.get("premium", 0),
                                                "policy_number": carrier.get(
                                                    "policy_number", ""
                                                ),
                                            }
                                        ]
                                    st.rerun()
                            with col_del:
                                if st.button(
                                    "🗑️",
                                    key=f"del_carrier_{original_idx}_{cidx}",
                                    use_container_width=True,
                                ):
                                    carriers_to_delete.append(cidx)

                            # Add Single Policy Number toggle if carrier has multiple RBEs
                            if carrier.get("has_multiple_rbes", False):
                                col_single_policy, _ = st.columns([1, 1])
                                with col_single_policy:
                                    single_policy = carrier.get(
                                        "single_policy_number", False
                                    )
                                    if st.button(
                                        (
                                            "🔗 Single Policy #"
                                            if not single_policy
                                            else "✅ Single Policy #"
                                        ),
                                        key=f"single_policy_{original_idx}_{cidx}",
                                        use_container_width=True,
                                        help="When enabled, uses the carrier's policy number for all RBEs in exports",
                                    ):
                                        carrier["single_policy_number"] = (
                                            not single_policy
                                        )
                                        st.rerun()

                            if not carrier.get("has_multiple_rbes", False):
                                # Simple carrier view (no RBE breakdown)
                                st.markdown("**Carrier Details:**")

                                c1, c2 = st.columns(2)
                                with c1:
                                    all_carriers = db.get_all_carriers()
                                    carrier_input = st.selectbox(
                                        "Carrier Name",
                                        options=[""] + all_carriers,
                                        index=(
                                            0
                                            if carrier.get("carrier_name", "") == ""
                                            else (
                                                all_carriers.index(
                                                    carrier.get("carrier_name", "")
                                                )
                                                + 1
                                                if carrier.get("carrier_name", "")
                                                in all_carriers
                                                else 0
                                            )
                                        ),
                                        key=f"carrier_name_{original_idx}_{cidx}",
                                    )
                                    if carrier_input == "":
                                        carrier_input = st.text_input(
                                            "Or type new carrier",
                                            value=carrier.get("carrier_name", ""),
                                            key=f"carrier_name_text_{original_idx}_{cidx}",
                                        )
                                    carrier["carrier_name"] = carrier_input

                                with c2:
                                    # Unified share percentage (improvement #1)
                                    carrier["share"] = (
                                        st.number_input(
                                            "Share %",
                                            value=float(carrier.get("share", 0) * 100),
                                            key=f"carrier_share_{original_idx}_{cidx}",
                                            format="%.2f",
                                            min_value=0.0,
                                            max_value=100.0,
                                        )
                                        / 100
                                    )

                                c3, c4 = st.columns(2)
                                with c3:
                                    carrier["premium"] = st.number_input(
                                        "Premium ($)",
                                        value=float(carrier.get("premium", 0)),
                                        key=f"carrier_prem_{original_idx}_{cidx}",
                                        format="%.0f",
                                        min_value=0.0,
                                    )

                                with c4:
                                    carrier["policy_number"] = st.text_input(
                                        "Policy #",
                                        value=carrier.get("policy_number", ""),
                                        key=f"carrier_policy_{original_idx}_{cidx}",
                                    )

                                st.markdown("**Fees & Policy:**")
                                f1, f2 = st.columns(2)

                                with f1:
                                    carrier["carrier_fee"] = st.number_input(
                                        "Carrier Fee ($)",
                                        value=float(carrier.get("carrier_fee", 0)),
                                        key=f"carrier_cfee_{original_idx}_{cidx}",
                                        format="%.2f",
                                        min_value=0.0,
                                    )

                                with f2:
                                    carrier["surplus_fee"] = st.number_input(
                                        "Surplus Fee ($)",
                                        value=float(carrier.get("surplus_fee", 0)),
                                        key=f"carrier_sfee_{original_idx}_{cidx}",
                                        format="%.2f",
                                        min_value=0.0,
                                    )
                            else:
                                # Multiple RBE view
                                st.markdown("**Carrier Info:**")

                                c1, c2 = st.columns(2)
                                with c1:
                                    # Ensure carrier_name is set
                                    if not carrier.get("carrier_name"):
                                        all_carriers = db.get_all_carriers()
                                        carrier_input = st.selectbox(
                                            "Carrier Name",
                                            options=[""] + all_carriers,
                                            key=f"carrier_name_multi_{original_idx}_{cidx}",
                                        )
                                        if carrier_input == "":
                                            carrier_input = st.text_input(
                                                "Or type new carrier",
                                                value="",
                                                key=f"carrier_name_text_multi_{original_idx}_{cidx}",
                                            )
                                        carrier["carrier_name"] = carrier_input
                                    else:
                                        st.markdown(
                                            f"**{carrier.get('carrier_name')}**"
                                        )

                                with c2:
                                    # Unified share percentage (improvement #1)
                                    carrier["share"] = (
                                        st.number_input(
                                            "Share %",
                                            value=float(carrier.get("share", 0) * 100),
                                            key=f"carrier_share_multi_{original_idx}_{cidx}",
                                            format="%.2f",
                                            min_value=0.0,
                                            max_value=100.0,
                                            help="Carrier's total participation in the layer",
                                        )
                                        / 100
                                    )

                                st.markdown("**Carrier Fees:**")
                                fee_col1, fee_col2 = st.columns(2)

                                with fee_col1:
                                    carrier["carrier_fee"] = st.number_input(
                                        "Carrier Fee ($)",
                                        value=float(carrier.get("carrier_fee", 0)),
                                        key=f"carrier_fee_{original_idx}_{cidx}",
                                        format="%.2f",
                                        min_value=0.0,
                                    )

                                with fee_col2:
                                    carrier["surplus_fee"] = st.number_input(
                                        "Surplus Lines Fee ($)",
                                        value=float(carrier.get("surplus_fee", 0)),
                                        key=f"surplus_fee_{original_idx}_{cidx}",
                                        format="%.2f",
                                        min_value=0.0,
                                    )

                                st.markdown("---")

                                # Add RBE button
                                if st.button(
                                    f"➕ Add RBE",
                                    key=f"add_rbe_{original_idx}_{cidx}",
                                    use_container_width=True,
                                ):
                                    if "rbes" not in carrier:
                                        carrier["rbes"] = []
                                    carrier["rbes"].append(
                                        {
                                            "rbe": "",
                                            "share": 0.0,
                                            "premium": 0,
                                            "policy_number": "",
                                        }
                                    )
                                    st.rerun()

                                st.markdown("**📋 Risk Bearing Entities:**")
                                carrier_share = carrier.get("share", 0)
                                st.caption(
                                    f"💡 RBE Share % is of carrier's {carrier_share*100:.1f}% layer participation (RBE shares should sum to 100%)"
                                )

                                # Ensure rbes list exists
                                if "rbes" not in carrier:
                                    carrier["rbes"] = []

                                if carrier["rbes"]:
                                    # Header row for RBEs
                                    h1, h2, h3, h4, h5 = st.columns(
                                        [2.5, 1, 1.2, 1.5, 0.4]
                                    )
                                    h1.markdown("**RBE Name**")
                                    h2.markdown("**RBE Share %**")
                                    h3.markdown("**Premium**")
                                    h4.markdown("**Policy #**")
                                    h5.markdown("")

                                    rbes_to_delete = []

                                    # Display each RBE
                                    for ridx, rbe in enumerate(carrier["rbes"]):
                                        r1, r2, r3, r4, r5 = st.columns(
                                            [2.5, 1, 1.2, 1.5, 0.4]
                                        )

                                        with r1:
                                            rbe["rbe"] = st.text_input(
                                                "RBE",
                                                value=rbe.get("rbe", ""),
                                                key=f"rbe_{original_idx}_{cidx}_{ridx}",
                                                label_visibility="collapsed",
                                                placeholder="Risk Bearing Entity",
                                            )

                                        with r2:
                                            rbe["share"] = (
                                                st.number_input(
                                                    "Share %",
                                                    value=float(
                                                        rbe.get("share", 0) * 100
                                                    ),
                                                    key=f"rbe_share_{original_idx}_{cidx}_{ridx}",
                                                    format="%.2f",
                                                    label_visibility="collapsed",
                                                    min_value=0.0,
                                                    max_value=100.0,
                                                    help=f"% of carrier's {carrier_share*100:.1f}% share",
                                                )
                                                / 100
                                            )

                                        with r3:
                                            rbe["premium"] = st.number_input(
                                                "Premium",
                                                value=float(rbe.get("premium", 0)),
                                                key=f"rbe_prem_{original_idx}_{cidx}_{ridx}",
                                                format="%.0f",
                                                label_visibility="collapsed",
                                                min_value=0.0,
                                            )

                                        with r4:
                                            rbe["policy_number"] = st.text_input(
                                                "Policy #",
                                                value=rbe.get("policy_number", ""),
                                                key=f"rbe_policy_{original_idx}_{cidx}_{ridx}",
                                                label_visibility="collapsed",
                                            )

                                        with r5:
                                            if st.button(
                                                "🗑️",
                                                key=f"del_rbe_{original_idx}_{cidx}_{ridx}",
                                            ):
                                                rbes_to_delete.append(ridx)

                                    # Delete RBEs
                                    for ridx in sorted(rbes_to_delete, reverse=True):
                                        carrier["rbes"].pop(ridx)
                                    if rbes_to_delete:
                                        st.rerun()

                                    # Show RBE share validation
                                    rbe_total = sum(
                                        rbe.get("share", 0) for rbe in carrier["rbes"]
                                    )
                                    if abs(rbe_total - 1.0) > 0.01:
                                        st.warning(
                                            f"⚠️ {carrier.get('carrier_name', 'Carrier')} RBE shares: {rbe_total*100:.1f}% (should be 100%)"
                                        )
                                    else:
                                        st.success(
                                            f"✅ {carrier.get('carrier_name', 'Carrier')} RBE shares: {rbe_total*100:.1f}%"
                                        )

                            st.markdown("</div>", unsafe_allow_html=True)
                            st.markdown("")

                            # Delete carriers
                        for cidx in sorted(carriers_to_delete, reverse=True):
                            layer["carriers"].pop(cidx)
                        if carriers_to_delete:
                            st.rerun()

                        # Total layer share validation
                        if layer["carriers"]:
                            total_layer_share = 0

                            for carrier in layer["carriers"]:
                                # Use unified share for all carriers
                                carrier_share = carrier.get("share", 0)
                                total_layer_share += carrier_share

                            # Display layer total validation
                            if abs(total_layer_share - 1.0) > 0.01:
                                st.warning(
                                    f"⚠️ Total layer share: {total_layer_share*100:.1f}% (should be 100%)"
                                )
                            else:
                                st.success(
                                    f"✅ Total layer share: {total_layer_share*100:.1f}%"
                                )

                for idx in sorted(layers_to_delete, reverse=True):
                    edited_program["layers"].pop(idx)
                if layers_to_delete:
                    st.rerun()

                # Save buttons
                st.markdown("---")
                col1, col2 = st.columns(2)

                with col1:
                    if st.button(
                        "💾 Save Changes", type="primary", use_container_width=True
                    ):
                        valid = True
                        validation_errors = []

                        for lidx, layer in enumerate(edited_program["layers"]):
                            if layer["carriers"]:
                                # Validate total layer share
                                total_layer_share = 0

                                for carrier in layer["carriers"]:
                                    carrier_name = carrier.get(
                                        "carrier_name", "Unknown"
                                    )

                                    # Use unified share for all carriers
                                    carrier_share = carrier.get("share", 0)
                                    total_layer_share += carrier_share

                                    if carrier.get("has_multiple_rbes", False):
                                        # Validate RBE shares sum to 100%
                                        rbe_total = sum(
                                            rbe.get("share", 0)
                                            for rbe in carrier.get("rbes", [])
                                        )
                                        if abs(rbe_total - 1.0) > 0.01:
                                            valid = False
                                            validation_errors.append(
                                                f"Layer {lidx + 1} - {carrier_name}: RBE shares sum to {rbe_total*100:.1f}% (should be 100%)"
                                            )

                                if abs(total_layer_share - 1.0) > 0.01:
                                    valid = False
                                    validation_errors.append(
                                        f"Layer {lidx + 1}: total layer share is {total_layer_share*100:.1f}% (should be 100%)"
                                    )

                        if valid:
                            # Make sure carrier_percent is set to share for consistency
                            for layer in edited_program["layers"]:
                                for carrier in layer.get("carriers", []):
                                    carrier["carrier_percent"] = carrier.get("share", 0)

                            db.save_program(selected_account_id, edited_program)
                            st.success("✅ Changes saved successfully!")
                            del st.session_state.edited_program
                            st.rerun()
                        else:
                            st.error("❌ Cannot save - validation errors:")
                            for err in validation_errors:
                                st.error(f"   {err}")

                with col2:
                    if st.button("🔄 Reset Changes", use_container_width=True):
                        del st.session_state.edited_program
                        st.rerun()

                display_program = edited_program
            else:
                if "edited_program" in st.session_state:
                    del st.session_state.edited_program
                display_program = program

            # Display Program
            st.markdown("---")
            st.markdown("### 📊 Program Structure Summary")

            if not display_program["layers"]:
                st.info("📝 No layers yet. Enable Build Mode to add layers.")
            else:
                # Tabular view by carrier
                sorted_layers = sorted(
                    display_program["layers"], key=lambda x: x.get("attachment", 0)
                )

                # Create tables for each layer (improvement #6)
                for idx, layer in enumerate(sorted_layers):
                    layer_title = format_layer_title(layer, idx)
                    st.markdown(f"**{layer_title}**")

                    # Create a table for all carriers in the layer
                    carriers = layer.get("carriers", [])
                    if carriers:
                        # Create dataframe for the layer's carriers
                        layer_data = []

                        for carrier in carriers:
                            carrier_name = carrier.get("carrier_name", "Unknown")
                            share = carrier.get("share", 0)
                            premium = carrier.get("premium", 0)
                            policy = carrier.get("policy_number", "")
                            carrier_fee = carrier.get("carrier_fee", 0)
                            surplus_fee = carrier.get("surplus_fee", 0)
                            total_fees = carrier_fee + surplus_fee

                            if carrier.get("has_multiple_rbes", False):
                                # Determine policy number display
                                if not carrier.get(
                                    "single_policy_number", False
                                ) and carrier.get("rbes", []):
                                    policy = "Multiple"

                                # For carriers with multiple RBEs, show as parent row
                                layer_data.append(
                                    {
                                        "Carrier": f"🏢 {carrier_name}",
                                        "Share %": f"{share*100:.1f}%",
                                        "Premium ($)": f"${premium:,.0f}",
                                        "Policy #": policy,
                                        "Fees ($)": f"${total_fees:,.2f}",
                                    }
                                )

                                # Add RBE rows indented
                                for rbe in carrier.get("rbes", []):
                                    rbe_name = rbe.get("rbe", "")
                                    rbe_share = rbe.get("share", 0)
                                    layer_share = rbe_share * share
                                    rbe_premium = rbe.get("premium", 0)

                                    # Use carrier policy number if single policy is enabled
                                    rbe_policy = rbe.get("policy_number", "")
                                    if carrier.get("single_policy_number", False):
                                        rbe_policy = policy

                                    layer_data.append(
                                        {
                                            "Carrier": f"    ↳ {rbe_name}",
                                            "Share %": f"{rbe_share*100:.1f}% of carrier ({layer_share*100:.2f}% of layer)",
                                            "Premium ($)": f"${rbe_premium:,.0f}",
                                            "Policy #": rbe_policy,
                                            "Fees ($)": "",
                                        }
                                    )
                            else:
                                # Simple carrier row
                                layer_data.append(
                                    {
                                        "Carrier": f"🏢 {carrier_name}",
                                        "Share %": f"{share*100:.1f}%",
                                        "Premium ($)": f"${premium:,.0f}",
                                        "Policy #": policy,
                                        "Fees ($)": f"${total_fees:,.2f}",
                                    }
                                )

                        # Display the layer table
                        df = pd.DataFrame(layer_data)
                        st.dataframe(df, use_container_width=True, hide_index=True)
                    else:
                        st.info(f"No carriers in {layer_title}")

                    st.markdown("---")

                # Mudmap visualization
                st.markdown("### 🗺️ Visual Program Structure (Mud Map)")
                st.caption(
                    "Each layer is shown as a horizontal band. Within each band, carriers are shown as blocks "
                    "sized proportionally to their share. Hover over blocks for details."
                )

                color_map = {}
                color_pool = [
                    "#A8D5BA",
                    "#FFD59E",
                    "#FFB5B5",
                    "#99CCFF",
                    "#F5CBA7",
                    "#C39BD3",
                    "#AED6F1",
                    "#A3E4D7",
                    "#F9E79F",
                    "#F1948A",
                    "#D7BDE2",
                    "#A9DFBF",
                    "#FAD7A0",
                    "#F5B7B1",
                    "#AED6F1",
                ]
                color_idx = 0

                fig = go.Figure()
                current_y = 0

                # Header Height adjusted for visibility
                header_height = 0.25

                layers_sorted = sorted(
                    display_program["layers"], key=lambda x: x.get("attachment", 0)
                )

                # --- CALCULATE VARIABLE LAYER HEIGHTS ---
                layer_heights = []
                # Minimum width enforcement (20%)
                MIN_VISUAL_WIDTH = 0.20

                for layer in layers_sorted:
                    carriers = layer.get("carriers", [])
                    max_lines = 1  # Minimum 1 line of text

                    if carriers:
                        # Normalize widths
                        visual_shares = [
                            max(c.get("share", 0), MIN_VISUAL_WIDTH) for c in carriers
                        ]
                        total_visual = sum(visual_shares)

                        normalized_shares = (
                            [s / total_visual for s in visual_shares]
                            if total_visual > 0
                            else [0] * len(carriers)
                        )

                        # Calculate text height for each carrier
                        for i, carrier in enumerate(carriers):
                            width = normalized_shares[i]
                            carrier_name = carrier.get("carrier_name", "Unknown")

                            # Estimate lines: 1600px width, 14px font ~ 10px char width
                            px_width = width * 1600
                            chars_per_line = max(10, int(px_width / 10))

                            wrapped_lines = textwrap.wrap(
                                carrier_name, width=chars_per_line
                            )
                            num_lines = (
                                len(wrapped_lines) + 3
                            )  # Name + Share + Premium + Spacing
                            if num_lines > max_lines:
                                max_lines = num_lines

                    # Store required height for this layer
                    layer_heights.append(1.0 + (max_lines * 0.25))

                # --- DRAWING LOOP ---
                for idx, layer in enumerate(layers_sorted):
                    layer_name = format_layer_title(layer, idx)
                    band_height = layer_heights[idx]

                    # Header Bar
                    fig.add_shape(
                        type="rect",
                        x0=0,
                        x1=1,
                        y0=current_y + band_height,
                        y1=current_y + band_height + header_height,
                        fillcolor="#D6DBDF",
                        line=dict(width=0),
                    )
                    fig.add_annotation(
                        x=0.5,
                        y=current_y + band_height + header_height / 2,
                        text=f"<b>{layer_name}</b>",
                        showarrow=False,
                        font=dict(size=16, family="Arial", color="black"),
                    )

                    # --- DRAW BOXES ---
                    carriers = layer.get("carriers", [])
                    visual_shares = [
                        max(c.get("share", 0), MIN_VISUAL_WIDTH) for c in carriers
                    ]
                    total_visual = sum(visual_shares)
                    normalized_shares = (
                        [s / total_visual for s in visual_shares]
                        if total_visual > 0
                        else [0] * len(carriers)
                    )

                    x0 = 0
                    for i, carrier in enumerate(carriers):
                        width = normalized_shares[i]
                        carrier_name = carrier.get("carrier_name", "Unknown")
                        carrier_share = carrier.get("share", 0)

                        if carrier_name not in color_map:
                            color_map[carrier_name] = color_pool[
                                color_idx % len(color_pool)
                            ]
                            color_idx += 1
                        fillcolor = color_map[carrier_name]

                        # Calculate total premium
                        if carrier.get("has_multiple_rbes", False):
                            total_premium = sum(
                                rbe.get("premium", 0) for rbe in carrier.get("rbes", [])
                            )
                        else:
                            total_premium = carrier.get("premium", 0)

                        # Draw Box
                        fig.add_shape(
                            type="rect",
                            x0=x0,
                            x1=x0 + width,
                            y0=current_y,
                            y1=current_y + band_height,
                            fillcolor=fillcolor,
                            line=dict(width=1, color="white"),
                        )

                        # Text Logic
                        policy_display = carrier.get("policy_number", "")
                        if carrier.get("has_multiple_rbes", False) and not carrier.get(
                            "single_policy_number", False
                        ):
                            policy_display = "Multiple"

                        hover_text = (
                            f"<b>{carrier_name}</b><br>"
                            + f"Share: {carrier_share*100:.2f}%<br>"
                            + f"Premium: ${total_premium:,.0f}<br>"
                            + f"Policy: {policy_display}"
                        )

                        # 1600px width scaling
                        px_width = width * 1600
                        char_capacity = max(10, int(px_width / 10))
                        wrapped_name = "<br>".join(
                            textwrap.wrap(carrier_name, width=char_capacity)
                        )

                        display_text = f"<b>{wrapped_name}</b><br>{carrier_share*100:.1f}%<br>${total_premium:,.0f}"

                        # Static Text (Horizontal)
                        fig.add_annotation(
                            x=x0 + width / 2,
                            y=current_y + band_height / 2,
                            text=display_text,
                            showarrow=False,
                            font=dict(size=14, family="Arial", color="black"),
                        )

                        # Hover Layer
                        fig.add_trace(
                            go.Scatter(
                                x=[x0 + width / 2],
                                y=[current_y + band_height / 2],
                                mode="markers",
                                marker=dict(size=1, opacity=0),
                                hoverinfo="text",
                                hovertext=[hover_text],
                                showlegend=False,
                            )
                        )

                        x0 += width

                    # Side Label (Left Axis)
                    attach = layer.get("attachment", 0)
                    limit = layer.get("limit", 0)

                    if layer.get("is_primary"):
                        side_label = f"Primary<br>${limit:,.0f}"
                    else:
                        # NEW FORMAT: Limit on top, Attachment on bottom
                        side_label = f"${limit:,.0f}<br>xs ${attach:,.0f}"

                    fig.add_annotation(
                        x=-0.01,  # moved closer to bars
                        y=current_y + band_height / 2,
                        text=f"<b>{side_label}</b>",
                        showarrow=False,
                        xanchor="right",
                        font=dict(size=14, family="Arial", color="black"),
                    )

                    current_y += band_height + header_height + 0.1

                fig.update_xaxes(range=[-0.35, 1], visible=False)
                fig.update_yaxes(visible=False)

                # Aggressive height scaling
                total_chart_height = max(900, sum(layer_heights) * 120 + 300)

                fig.update_layout(
                    autosize=True,
                    height=total_chart_height,
                    margin=dict(
                        l=50, r=80, t=60, b=40
                    ),  # Reduced left margin since labels are inside range
                    plot_bgcolor="white",
                )

                st.plotly_chart(fig, use_container_width=True, theme=None)

                # Export buttons
                st.markdown("---")
                st.markdown("### 📤 Export Options")
                col1, col2, col3 = st.columns(3)

                with col1:
                    excel_file = export_program_to_excel(display_program, None)
                    st.download_button(
                        label="📊 Download Excel",
                        data=excel_file,
                        file_name=f"{display_program['account']}_program.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                with col2:
                    word_file = export_program_to_word(display_program, None)
                    st.download_button(
                        label="📄 Download Word",
                        data=word_file,
                        file_name=f"{display_program['account']}_program.docx",
                        mime="application/vnd.openxmlformats-officedocument.docx",
                        use_container_width=True,
                    )

                with col3:
                    try:
                        # Export visualization as PDF
                        pdf_bytes = fig.to_image(
                            format="pdf", width=1600, height=total_chart_height
                        )
                        st.download_button(
                            label="📑 Download PDF",
                            data=pdf_bytes,
                            file_name=f"{display_program['account']}_visual_structure.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(
                            f"PDF export requires kaleido package. Install with: pip install kaleido"
                        )

# Carrier Library
elif menu == "📚 Carrier Library":
    st.markdown(
        "<h1 class='main-header'>📚 Carrier Library</h1>", unsafe_allow_html=True
    )

    st.markdown(
        "Manage your carrier database for quick selection when building programs."
    )

    with styled_expander("➕ Add New Carrier", expanded=False):
        col1, col2 = st.columns([3, 1])
        with col1:
            new_carrier_name = st.text_input("Carrier Name")
        with col2:
            st.write("")
            st.write("")
            if st.button("Add Carrier", type="primary", use_container_width=True):
                if new_carrier_name:
                    if db.add_carrier(new_carrier_name):
                        st.success(f"✅ Carrier '{new_carrier_name}' added!")
                        st.rerun()
                    else:
                        st.error("❌ Carrier already exists!")
                else:
                    st.warning("⚠️ Please enter a carrier name")

    st.markdown("---")
    st.markdown("### 📋 All Carriers")

    carriers = db.get_all_carriers()
    if carriers:
        st.info(
            "💡 **Note:** Carrier fees and surplus lines fees are set per carrier in Build Program."
        )

        for carrier in carriers:
            col1, col2 = st.columns([4, 1])
            with col1:
                st.write(f"**{carrier}**")
            with col2:
                if st.button(
                    "🗑️ Delete",
                    key=f"del_carrier_{carrier}",
                    use_container_width=True,
                ):
                    db.delete_carrier(carrier)
                    st.rerun()
    else:
        st.info("📝 No carriers registered yet. Add your first carrier above!")

# Settings
elif menu == "⚙️ Settings":
    st.markdown("<h1 class='main-header'>⚙️ Settings</h1>", unsafe_allow_html=True)

    st.markdown("### 🗄️ Database Management")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Backup Database")
        if st.button("💾 Create Backup", use_container_width=True):
            import shutil
            from datetime import datetime

            backup_path = (
                f"data/insurance_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            )
            shutil.copy2(db.db_path, backup_path)
            st.success(f"✅ Backup created: {backup_path}")

    with col2:
        st.markdown("#### Database Info")
        accounts = db.get_all_accounts()
        carriers = db.get_all_carriers()
        st.info(f"**Accounts:** {len(accounts)}\n\n**Carriers:** {len(carriers)}")

    st.markdown("---")
    st.markdown("### 📊 Application Info")
    st.info(
        """
    **Insurance Layer Builder v1.0.0**
    
    Standalone application for managing insurance account programs with layered structures.
    
    Features:
    - Multi-account management
    - Layer and carrier participation tracking
    - Unified share percentage for both visual representation and actual participation
    - Optional Multiple Risk Bearing Entities (RBEs) per carrier
    - RBE shares are % of carrier's participation (not layer)
    - Carrier-level fees for simplified management
    - Visual mudmap representation with carrier color grouping
    - Excel and Word export with full details
    - Carrier library for quick selection
    - Local SQLite database storage
    """
    )

# Check if we need to navigate from dashboard
if "current_menu" in st.session_state:
    if st.session_state.current_menu != menu:
        st.rerun()
