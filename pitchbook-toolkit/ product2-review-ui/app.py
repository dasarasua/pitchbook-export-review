import streamlit as st
import pandas as pd

# Configuration
EXCEL_FILE = "companies.xlsx"
WORKING_SHEET = "Working"
APPROVED_SHEET = "Approved"
DISAPPROVED_SHEET = "Disapproved"

# CSS
st.markdown("""
    <style>
    div.stForm {
        border: none !important;
        box-shadow: none !important;
        padding: 0px !important;
        margin: 0px !important;
    }
    textarea {
        margin-bottom: 5px !important;
    }

    /* Force buttons by title */
    button[title="Approve"], button[title="✅ Approved"] {
        background-color: green !important;
        color: white !important;
    }
    button[title="Disapprove"], button[title="❌ Disapproved"] {
        background-color: red !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# Load Excel data
df_working = pd.read_excel(EXCEL_FILE, sheet_name=WORKING_SHEET)

try:
    df_approved = pd.read_excel(EXCEL_FILE, sheet_name=APPROVED_SHEET)
    if "Grade" not in df_approved.columns:
        df_approved["Grade"] = 0.0
    approved_company_names = df_approved["Company Name"].tolist()
    approved_comments_dict = dict(zip(df_approved["Company Name"], df_approved["Comments"]))
    approved_grades_dict = dict(zip(df_approved["Company Name"], df_approved["Grade"]))
except:
    df_approved = pd.DataFrame(columns=["Company Name", "Grade", "Comments"] + df_working.columns.tolist())
    approved_company_names = []
    approved_comments_dict = {}
    approved_grades_dict = {}

try:
    df_disapproved = pd.read_excel(EXCEL_FILE, sheet_name=DISAPPROVED_SHEET)
    if "Grade" not in df_disapproved.columns:
        df_disapproved["Grade"] = 0.0
    disapproved_company_names = df_disapproved["Company Name"].tolist()
    disapproved_comments_dict = dict(zip(df_disapproved["Company Name"], df_disapproved["Comments"]))
    disapproved_grades_dict = dict(zip(df_disapproved["Company Name"], df_disapproved["Grade"]))
except:
    df_disapproved = pd.DataFrame(columns=["Company Name", "Grade", "Comments"] + df_working.columns.tolist())
    disapproved_company_names = []
    disapproved_comments_dict = {}
    disapproved_grades_dict = {}

# State
if "current_index" not in st.session_state:
    st.session_state.current_index = 0

# Sidebar
filter_option = st.sidebar.radio("Filter rows:", ("All", "Approved", "Disapproved", "Not evaluated"))

if filter_option == "All":
    df_filtered = df_working
elif filter_option == "Approved":
    df_filtered = df_working[df_working["Company Name"].isin(approved_company_names)]
elif filter_option == "Disapproved":
    df_filtered = df_working[df_working["Company Name"].isin(disapproved_company_names)]
elif filter_option == "Not evaluated":
    df_filtered = df_working[~df_working["Company Name"].isin(approved_company_names + disapproved_company_names)]

if st.session_state.current_index >= len(df_filtered):
    st.session_state.current_index = 0

# Arrows
top_prev, top_spacer, top_next = st.columns([2, 6, 2])
with top_prev:
    prev_clicked = st.button("◀️", use_container_width=True)
with top_next:
    next_clicked = st.button("▶️", use_container_width=True)

if prev_clicked and st.session_state.current_index > 0:
    st.session_state.current_index -= 1
    st.rerun()
if next_clicked and st.session_state.current_index < len(df_filtered) - 1:
    st.session_state.current_index += 1
    st.rerun()

# Show row
if len(df_filtered) > 0:
    row = df_filtered.iloc[st.session_state.current_index]

    st.markdown(
        f"<div style='color:gray; opacity:0.6; font-size:13px;'>Row {st.session_state.current_index + 1} of {len(df_filtered)}</div>",
        unsafe_allow_html=True
    )

    is_approved = row["Company Name"] in approved_company_names
    is_disapproved = row["Company Name"] in disapproved_company_names

    if is_approved:
        title_color = "green"
    elif is_disapproved:
        title_color = "red"
    else:
        title_color = "black"

    col1, col2 = st.columns([4, 1])
    with col1:
        st.markdown(
            f"<h2 style='text-decoration: underline; font-size:29px; color:{title_color};'>{row['Company Name']} "
            f"<span style='font-size:17px; font-weight:normal; color:gray;'> - {row['Primary Industry Sector']}</span></h2>",
            unsafe_allow_html=True
        )

        website = row['Website']
        if pd.notnull(website):
            if not website.startswith("http"):
                website = "https://" + website
            st.markdown(
                f"<div style='font-size:16px; margin-bottom:4px;'><a href='{website}' target='_blank' style='color:blue;'>{website}</a></div>",
                unsafe_allow_html=True
            )

        if pd.notnull(row['Employees']):
            st.markdown(
                f"<div style='font-size:15px; margin-bottom:4px;'>{int(row['Employees'])} Employees</div>",
                unsafe_allow_html=True
            )
    with col2:
        st.markdown(
            f"<div style='font-size:15px; text-align:right; color:gray;'>Year Founded: {int(row['Year Founded']) if pd.notnull(row['Year Founded']) else ''}</div>",
            unsafe_allow_html=True
        )

    st.markdown("---")

    st.markdown(
        f"<div style='font-size:16px; margin-bottom:4px;'><b>Description:</b> {row['Description']}</div>",
        unsafe_allow_html=True
    )
    st.markdown(
        f"<div style='font-size:16px; margin-top:4px;'><b>Financing Status Note:</b> {row['Financing Status Note']}</div>",
        unsafe_allow_html=True
    )

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"<div style='font-size:15px;'><b>All Industries:</b> {row['All Industries']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Verticals:</b> {row['Verticals']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Active Investors:</b> {row['Active Investors']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Business Status:</b> {row['Business Status']}</div>", unsafe_allow_html=True)

    with col2:
        st.markdown(f"<div style='font-size:15px;'><b>Company Financing Status:</b> {row['Company Financing Status']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Last Financing Size:</b> {row['Last Financing Size']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Last Financing Date:</b> {row['Last Financing Date']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Last Financing Deal Type:</b> {row['Last Financing Deal Type']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>Last Known Valuation:</b> {row['Last Known Valuation']}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:15px;'><b>First Financing Size:</b> {row['First Financing Size']}</div>", unsafe_allow_html=True)

    st.markdown("---")

    # GRADE
    existing_grade = approved_grades_dict.get(row["Company Name"], 0.0)
    if not existing_grade:
        existing_grade = disapproved_grades_dict.get(row["Company Name"], 0.0)
    existing_grade = float(existing_grade)

    # COMMENTS
    existing_comment = approved_comments_dict.get(row["Company Name"], "")
    if not existing_comment:
        existing_comment = disapproved_comments_dict.get(row["Company Name"], "")
    if pd.isna(existing_comment):
        existing_comment = ""

    # Initialize session state for inputs if not exists
    if "grade_input" not in st.session_state:
        st.session_state.grade_input = 0.0
    if "comment_input" not in st.session_state:
        st.session_state.comment_input = ""

    # If current row changed, reset inputs to existing saved values
    if "last_row_company" not in st.session_state or st.session_state.last_row_company != row["Company Name"]:
        st.session_state.grade_input = existing_grade
        st.session_state.comment_input = existing_comment
        st.session_state.last_row_company = row["Company Name"]

    # INPUTS
    grade_input = st.number_input(
        "Grade",
        min_value=0.0,
        max_value=10.0,
        step=0.1,
        value=st.session_state.grade_input,
        key="grade_input"
    )
    comment_input = st.text_area(
        "Comments",
        value=st.session_state.comment_input,
        key="comment_input"
    )

    # BUTTONS: Approve + Disapprove — one form
    with st.form(key='evaluate_form'):
        col1, col2, col3, col4 = st.columns([2, 3, 3, 2])
        
        with col2:
            approve_button = st.form_submit_button(
                "✅ Approved" if is_approved else "Approve"
            )
        with col3:
            disapprove_button = st.form_submit_button(
                "❌ Disapproved" if is_disapproved else "Disapprove"
            )

    # Excel logic
    if approve_button:
        df_disapproved = df_disapproved[df_disapproved["Company Name"] != row["Company Name"]]
        df_approved = df_approved[df_approved["Company Name"] != row["Company Name"]]
        new_row = pd.DataFrame([[row["Company Name"], grade_input, comment_input] + row.tolist()],
                               columns=df_approved.columns)
        df_approved = pd.concat([df_approved, new_row], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            df_approved.to_excel(writer, sheet_name=APPROVED_SHEET, index=False)
            df_disapproved.to_excel(writer, sheet_name=DISAPPROVED_SHEET, index=False)

        st.session_state.clear()
        st.rerun()

    elif disapprove_button:
        df_approved = df_approved[df_approved["Company Name"] != row["Company Name"]]
        df_disapproved = df_disapproved[df_disapproved["Company Name"] != row["Company Name"]]
        new_row = pd.DataFrame([[row["Company Name"], grade_input, comment_input] + row.tolist()],
                               columns=df_disapproved.columns)
        df_disapproved = pd.concat([df_disapproved, new_row], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            df_approved.to_excel(writer, sheet_name=APPROVED_SHEET, index=False)
            df_disapproved.to_excel(writer, sheet_name=DISAPPROVED_SHEET, index=False)

        st.session_state.clear()
        st.rerun()

else:
    st.warning("No rows match your filter!")
