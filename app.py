import streamlit as st
import pandas as pd
import plotly.express as px
import json
import io
from datetime import datetime, timedelta

st.set_page_config(
    page_title="M365 Unified Audit Log Viewer",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("M365 Unified Audit Log Viewer")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Parsing CSV…")
def load_csv(file_bytes: bytes) -> pd.DataFrame:
    """Read uploaded CSV and parse the AuditData JSON column."""
    df = pd.read_csv(io.BytesIO(file_bytes))

    # Normalise column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Parse CreationDate
    if "CreationDate" in df.columns:
        df["CreationDate"] = pd.to_datetime(
            df["CreationDate"].str.strip(), utc=True, errors="coerce"
        )

    # Parse AuditData JSON
    parsed_records = []
    for raw in df["AuditData"]:
        try:
            parsed_records.append(json.loads(raw))
        except (json.JSONDecodeError, TypeError):
            parsed_records.append({})

    df["_audit_parsed"] = parsed_records

    # Extract commonly useful fields from AuditData
    df["Workload"] = df["_audit_parsed"].apply(lambda d: d.get("Workload", ""))
    df["ResultStatus"] = df["_audit_parsed"].apply(lambda d: d.get("ResultStatus", ""))
    df["ObjectId"] = df["_audit_parsed"].apply(lambda d: d.get("ObjectId", ""))
    df["ClientIP"] = df["_audit_parsed"].apply(lambda d: d.get("ClientIP", d.get("ClientIPAddress", "")))

    return df


def try_parse_json(value):
    """Recursively try to parse JSON strings into Python objects."""
    if not isinstance(value, str):
        return value
    stripped = value.strip()
    if not stripped or (not stripped.startswith('{') and not stripped.startswith('[')):
        return value
    try:
        parsed = json.loads(stripped)
        # Recursively resolve nested JSON strings in the result
        if isinstance(parsed, dict):
            return {k: try_parse_json(v) for k, v in parsed.items()}
        if isinstance(parsed, list):
            return [try_parse_json(item) for item in parsed]
        return parsed
    except (json.JSONDecodeError, TypeError):
        return value


def render_value(value, *, inline: bool = False) -> str:
    """Render a value as a readable string, pretty-printing nested JSON."""
    resolved = try_parse_json(value)
    if isinstance(resolved, (dict, list)):
        return json.dumps(resolved, indent=2)
    text = str(resolved) if resolved is not None else ""
    return text


def render_kv(label: str, value, *, skip_empty: bool = True):
    """Render a single key-value pair as markdown, skipping empty values."""
    if value is None and skip_empty:
        return
    s = str(value).strip() if value is not None else ""
    if skip_empty and s in ("", "None"):
        return
    st.markdown(f"**{label}:** {s}")


def render_dict_section(data: dict, title: str = "", *, expanded: bool = True, skip_keys: set = None):
    """Render a dict as an expander with key-value pairs."""
    skip_keys = skip_keys or set()
    items = {k: v for k, v in data.items() if k not in skip_keys and v is not None and str(v).strip() not in ("", "None")}
    if not items:
        return
    container = st.expander(title, expanded=expanded) if title else st
    with container:
        for k, v in items.items():
            if isinstance(v, (dict, list)):
                st.markdown(f"**{k}:**")
                st.json(v)
            else:
                st.markdown(f"**{k}:** {v}")


def render_extra_props(entries: list, *, use_key_value: bool = False):
    """Render ExtraProperties list (Teams-style Key/Value pairs or Name/Value)."""
    if not entries:
        return
    if use_key_value:
        rows = []
        for e in entries:
            rows.append({"Property": e.get("Key", e.get("Name", "")), "Value": e.get("Value", "")})
    else:
        rows = []
        for e in entries:
            rows.append({"Property": e.get("Name", e.get("Key", "")), "Value": e.get("Value", "")})
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_attendees(attendees: list):
    """Render Teams meeting/call attendees."""
    if not attendees:
        return
    rows = []
    for a in attendees:
        rows.append({
            "Name": a.get("DisplayName", ""),
            "UPN": a.get("UPN", ""),
            "Role": {0: "Attendee", 1: "Presenter", 2: "Organizer"}.get(a.get("Role"), str(a.get("Role", ""))),
            "Type": a.get("RecipientType", ""),
            "Organizer": a.get("IsOrganizer", ""),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_parameters(params: list):
    """Render Exchange cmdlet parameters."""
    if not params:
        return
    rows = []
    for p in params:
        val = render_value(p.get("Value", ""))
        rows.append({"Parameter": p.get("Name", ""), "Value": val})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_affected_items(items: list):
    """Render Exchange AffectedItems list."""
    if not items:
        return
    for i, item in enumerate(items):
        st.markdown(f"**Item {i + 1}**")
        for k, v in item.items():
            if isinstance(v, (dict, list)):
                st.markdown(f"**{k}:**")
                st.json(v)
            else:
                st.markdown(f"**{k}:** {v}")
        if i < len(items) - 1:
            st.divider()


def format_actor_target(entries: list) -> pd.DataFrame:
    """Turn Actor/Target list-of-dicts into a readable table."""
    type_map = {0: "Default", 1: "Name", 2: "ObjectId", 3: "PUID", 4: "SPN", 5: "UPN"}
    rows = []
    for e in entries:
        rows.append({
            "Value": e.get("ID", ""),
            "Type": type_map.get(e.get("Type"), str(e.get("Type", ""))),
        })
    return pd.DataFrame(rows)


def render_modified_props(entries: list):
    """Render ModifiedProperties with nested JSON resolved and text wrapping."""
    for e in entries:
        name = e.get("Name", "")
        old_raw = e.get("OldValue", "")
        new_raw = e.get("NewValue", "")
        old = render_value(old_raw)
        new = render_value(new_raw)

        st.markdown(f"**{name}**")
        col_old, col_new = st.columns(2)
        with col_old:
            st.caption("Old Value")
            if old.startswith('{') or old.startswith('['):
                st.code(old, language="json")
            elif old:
                st.text(old)
            else:
                st.text("(empty)")
        with col_new:
            st.caption("New Value")
            if new.startswith('{') or new.startswith('['):
                st.code(new, language="json")
            elif new:
                st.text(new)
            else:
                st.text("(empty)")
        st.divider()


def render_extended_props(entries: list):
    """Render ExtendedProperties with nested JSON resolved and text wrapping."""
    for e in entries:
        name = e.get("Name", "")
        val = render_value(e.get("Value", ""))

        st.markdown(f"**{name}**")
        if val.startswith('{') or val.startswith('['):
            st.code(val, language="json")
        elif val:
            st.text(val)
        else:
            st.text("(empty)")


# ---------------------------------------------------------------------------
# File Upload
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Upload a Unified Audit Log CSV",
    type=["csv"],
    help="Export from Microsoft Purview / Security & Compliance Center → Audit log search → Export results.",
)

if uploaded is None:
    st.info("Upload a CSV file to get started. The file should contain an **AuditData** column with JSON log entries.")
    st.stop()

df = load_csv(uploaded.getvalue())

if "AuditData" not in df.columns:
    st.error("The uploaded CSV does not contain an **AuditData** column. Please upload a valid Unified Audit Log export.")
    st.stop()

# ---------------------------------------------------------------------------
# Sidebar Filters
# ---------------------------------------------------------------------------

st.sidebar.header("Filters")

# Date range
min_date = df["CreationDate"].min()
max_date = df["CreationDate"].max()

if pd.notna(min_date) and pd.notna(max_date):
    date_range = st.sidebar.date_input(
        "Date range",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )
else:
    date_range = None

# Operation
all_operations = sorted(df["Operation"].dropna().unique().tolist())
selected_ops = st.sidebar.multiselect("Operation", all_operations, default=[])

# UserId
all_users = sorted(df["UserId"].dropna().unique().tolist())
selected_users = st.sidebar.multiselect("User", all_users, default=[])

# Workload
all_workloads = sorted(df["Workload"].dropna().unique().tolist())
selected_workloads = st.sidebar.multiselect("Workload", all_workloads, default=[])

# Free text search
search_text = st.sidebar.text_input("Search (all fields)", placeholder="e.g. user@domain.com")

# Apply filters
mask = pd.Series(True, index=df.index)

if date_range and len(date_range) == 2:
    start, end = date_range
    mask &= df["CreationDate"].dt.date >= start
    mask &= df["CreationDate"].dt.date <= end

if selected_ops:
    mask &= df["Operation"].isin(selected_ops)

if selected_users:
    mask &= df["UserId"].isin(selected_users)

if selected_workloads:
    mask &= df["Workload"].isin(selected_workloads)

if search_text:
    text_lower = search_text.lower()
    # Search across several string columns + raw AuditData
    text_cols = ["Operation", "UserId", "Workload", "ResultStatus", "ObjectId", "ClientIP", "AuditData"]
    text_mask = pd.Series(False, index=df.index)
    for col in text_cols:
        if col in df.columns:
            text_mask |= df[col].astype(str).str.lower().str.contains(text_lower, na=False)
    mask &= text_mask

filtered = df[mask].copy()

st.sidebar.markdown(f"**Showing {len(filtered):,}** of {len(df):,} records")

# ---------------------------------------------------------------------------
# Summary Dashboard
# ---------------------------------------------------------------------------

st.header("Summary")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Records", f"{len(filtered):,}")
col2.metric("Unique Users", filtered["UserId"].nunique())

if pd.notna(filtered["CreationDate"].min()) and pd.notna(filtered["CreationDate"].max()):
    span = filtered["CreationDate"].max() - filtered["CreationDate"].min()
    col3.metric("Time Span", f"{span.days}d {span.seconds // 3600}h")
else:
    col3.metric("Time Span", "N/A")

if len(filtered) > 0:
    top_op = filtered["Operation"].value_counts().idxmax()
    col4.metric("Top Operation", top_op)
else:
    col4.metric("Top Operation", "N/A")

# Charts
chart1, chart2, chart3 = st.columns(3)

with chart1:
    op_counts = filtered["Operation"].value_counts().reset_index()
    op_counts.columns = ["Operation", "Count"]
    fig = px.bar(op_counts, x="Count", y="Operation", orientation="h", title="Events by Operation")
    fig.update_layout(yaxis=dict(autorange="reversed"), height=350, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig, use_container_width=True)

with chart2:
    user_counts = filtered["UserId"].value_counts().head(15).reset_index()
    user_counts.columns = ["User", "Count"]
    fig = px.bar(user_counts, x="Count", y="User", orientation="h", title="Events by User (top 15)")
    fig.update_layout(yaxis=dict(autorange="reversed"), height=350, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig, use_container_width=True)

with chart3:
    wl_counts = filtered["Workload"].value_counts().reset_index()
    wl_counts.columns = ["Workload", "Count"]
    fig = px.bar(wl_counts, x="Count", y="Workload", orientation="h", title="Events by Workload")
    fig.update_layout(yaxis=dict(autorange="reversed"), height=350, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# Timeline View
# ---------------------------------------------------------------------------

st.header("Timeline")

tl_col1, tl_col2 = st.columns([1, 3])
with tl_col1:
    color_by = st.radio("Color by", ["Operation", "Workload"], horizontal=True)
    agg_mode = st.radio("Mode", ["Individual events", "Daily count"], horizontal=True)

with tl_col2:
    if len(filtered) > 0:
        if agg_mode == "Individual events":
            tl_df = filtered[["CreationDate", "Operation", "Workload", "UserId", "ObjectId"]].copy()
            tl_df = tl_df.sort_values("CreationDate")
            fig = px.scatter(
                tl_df,
                x="CreationDate",
                y=color_by,
                color=color_by,
                hover_data=["UserId", "ObjectId", "Operation"],
                title="Audit Events Over Time",
            )
            fig.update_traces(marker=dict(size=8, opacity=0.7))
            fig.update_layout(height=400, margin=dict(l=0, r=0, t=40, b=0), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            daily = filtered.copy()
            daily["Date"] = daily["CreationDate"].dt.date
            daily_counts = daily.groupby(["Date", color_by]).size().reset_index(name="Count")
            fig = px.bar(
                daily_counts,
                x="Date",
                y="Count",
                color=color_by,
                title="Daily Event Counts",
            )
            fig.update_layout(height=400, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No events match the current filters.")

# ---------------------------------------------------------------------------
# Log Table
# ---------------------------------------------------------------------------

st.header("Log Entries")

display_cols = ["CreationDate", "Operation", "UserId", "Workload", "ResultStatus", "ObjectId", "ClientIP"]
sorted_filtered = filtered[[c for c in display_cols + ["_audit_parsed"] if c in filtered.columns]].copy()
sorted_filtered = sorted_filtered.sort_values("CreationDate", ascending=False).reset_index(drop=True)
display_df = sorted_filtered[[c for c in display_cols if c in sorted_filtered.columns]]

# Show the table with row selection
selection = st.dataframe(
    display_df,
    use_container_width=True,
    height=400,
    column_config={
        "CreationDate": st.column_config.DatetimeColumn("Date/Time", format="YYYY-MM-DD HH:mm:ss"),
    },
    on_select="rerun",
    selection_mode="single-row",
)

# ---------------------------------------------------------------------------
# Detail Viewer — driven by table row selection
# ---------------------------------------------------------------------------

st.header("Entry Detail")

if len(sorted_filtered) == 0:
    st.info("No entries to display.")
    st.stop()

selected_rows = selection.selection.rows
if not selected_rows:
    st.info("Click a row in the table above to view its details.")
    st.stop()

selected_idx = selected_rows[0]
row = sorted_filtered.iloc[selected_idx]
audit = row["_audit_parsed"]

if not audit:
    st.warning("Could not parse AuditData for this entry.")
    st.stop()

# Top-level summary
detail_cols = st.columns(3)
detail_cols[0].markdown(f"**Operation:** {audit.get('Operation', 'N/A')}")
detail_cols[0].markdown(f"**Result:** {audit.get('ResultStatus', 'N/A')}")
detail_cols[1].markdown(f"**Workload:** {audit.get('Workload', 'N/A')}")
detail_cols[1].markdown(f"**User:** {audit.get('UserId', 'N/A')}")
detail_cols[2].markdown(f"**Object:** {audit.get('ObjectId', 'N/A')}")
detail_cols[2].markdown(f"**Time:** {audit.get('CreationTime', 'N/A')}")

workload = audit.get("Workload", "")
operation = audit.get("Operation", "")

# =========================================================================
# Azure Active Directory
# =========================================================================
if workload == "AzureActiveDirectory":
    # Login-specific fields
    if operation in ("UserLoggedIn", "UserLoginFailed"):
        with st.expander("Sign-In Details", expanded=True):
            render_kv("IP Address", audit.get("ActorIpAddress") or audit.get("ClientIP"))
            render_kv("Application ID", audit.get("ApplicationId"))
            # ExtendedProperties often has ResultStatusDetail, UserAgent, RequestType
            for ep in audit.get("ExtendedProperties", []):
                render_kv(ep.get("Name", ""), ep.get("Value", ""))
        if operation == "UserLoginFailed":
            with st.expander("Error Details", expanded=True):
                render_kv("Error Number", audit.get("ErrorNumber"))
                render_kv("Logon Error", audit.get("LogonError"))
        # Device Properties
        dev_props = audit.get("DeviceProperties", [])
        if dev_props:
            with st.expander("Device Properties", expanded=False):
                rows = [{"Property": d.get("Name", ""), "Value": d.get("Value", "")} for d in dev_props]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        # Extended Properties
        ext_props = audit.get("ExtendedProperties", [])
        if ext_props:
            with st.expander("Extended Properties", expanded=False):
                render_extended_props(ext_props)

    # Device Properties (non-login ops that may also have them)
    if operation not in ("UserLoggedIn", "UserLoginFailed"):
        dev_props = audit.get("DeviceProperties", [])
        if dev_props:
            with st.expander("Device Properties", expanded=False):
                rows = [{"Property": d.get("Name", ""), "Value": d.get("Value", "")} for d in dev_props]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Actor
    actors = audit.get("Actor", [])
    if actors:
        with st.expander("Actor", expanded=True):
            st.dataframe(format_actor_target(actors), use_container_width=True, hide_index=True)

    # Target
    targets = audit.get("Target", [])
    if targets:
        with st.expander("Target", expanded=True):
            st.dataframe(format_actor_target(targets), use_container_width=True, hide_index=True)

    # Modified Properties
    mod_props = audit.get("ModifiedProperties", [])
    if mod_props:
        with st.expander("Modified Properties", expanded=True):
            render_modified_props(mod_props)

# =========================================================================
# Exchange
# =========================================================================
elif workload == "Exchange":
    # Distinguish cmdlet ops (Set-Mailbox, etc.) from mailbox item ops
    is_cmdlet = operation.startswith(("Set-", "New-", "Remove-", "Add-", "Enable-", "Disable-", "Install-", "Update-", "Get-"))

    if is_cmdlet:
        # Cmdlet operation
        with st.expander("Cmdlet Details", expanded=True):
            render_kv("Cmdlet", operation)
            render_kv("Object", audit.get("ObjectId"))
            render_kv("Result", audit.get("ResultStatus"))
            render_kv("Organization", audit.get("OrganizationName"))
            render_kv("App Pool", audit.get("AppPoolName"))
            render_kv("Client Process", audit.get("ClientProcessName"))
            render_kv("Originating Server", (audit.get("OriginatingServer") or "").strip())

        # Parameters
        params = audit.get("Parameters", [])
        if params:
            with st.expander("Parameters", expanded=True):
                render_parameters(params)

        # Modified Properties (some cmdlets have these)
        mod_props = audit.get("ModifiedProperties", [])
        if mod_props:
            with st.expander("Modified Properties", expanded=True):
                render_modified_props(mod_props)

    else:
        # Mailbox item operations (Send, Create, SoftDelete, HardDelete, etc.)

        # Mailbox info
        mb_owner = audit.get("MailboxOwnerUPN")
        if mb_owner:
            with st.expander("Mailbox", expanded=True):
                render_kv("Mailbox Owner", mb_owner)
                render_kv("Mailbox GUID", audit.get("MailboxGuid"))
                render_kv("Organization", audit.get("OrganizationName"))
                render_kv("External Access", audit.get("ExternalAccess"))
                render_kv("Cross-Mailbox Operation", audit.get("CrossMailboxOperation"))

        # Item details
        item = audit.get("Item")
        if item and isinstance(item, dict):
            with st.expander("Item Details", expanded=True):
                render_kv("Subject", item.get("Subject"))
                render_kv("Size", f"{item.get('SizeInBytes', '')} bytes" if item.get("SizeInBytes") else None)
                render_kv("Attachments", item.get("Attachments"))
                render_kv("Internet Message ID", item.get("InternetMessageId"))
                render_kv("ID", item.get("Id"))
                render_kv("Immutable ID", item.get("ImmutableId"))
                # ParentFolder
                pf = item.get("ParentFolder")
                if pf and isinstance(pf, dict):
                    st.divider()
                    st.markdown("**Parent Folder**")
                    render_kv("Path", pf.get("Path"))
                    render_kv("Name", pf.get("Name"))
                    render_kv("Member Rights", pf.get("MemberRights"))
                    render_kv("Member UPN", pf.get("MemberUpn"))
                    render_kv("Member SID", pf.get("MemberSid"))

        # Affected Items (bulk operations like SoftDelete)
        affected = audit.get("AffectedItems", [])
        if affected:
            with st.expander(f"Affected Items ({len(affected)})", expanded=True):
                render_affected_items(affected)

        # Folder info
        folder = audit.get("Folder")
        if folder and isinstance(folder, dict):
            with st.expander("Folder", expanded=True):
                for k, v in folder.items():
                    render_kv(k, v)

        # Folders (plural, for some ops)
        folders = audit.get("Folders")
        if folders and isinstance(folders, list):
            with st.expander(f"Folders ({len(folders)})", expanded=False):
                for f_item in folders:
                    if isinstance(f_item, dict):
                        for k, v in f_item.items():
                            render_kv(k, v)
                        st.divider()

        # Destination folder (for moves)
        dest = audit.get("DestFolder")
        if dest and isinstance(dest, dict):
            with st.expander("Destination Folder", expanded=True):
                for k, v in dest.items():
                    render_kv(k, v)

        render_kv("Save to Sent Items", audit.get("SaveToSentItems"))

    # Client & Logon (common to all Exchange)
    cl_vals = {k: audit.get(k) for k in [
        "ClientIPAddress", "ClientInfoString", "ClientAppId", "AppId",
        "HostAppId", "AuthType", "LogonType", "InternalLogonType",
        "TokenType", "ClientRequestId", "DeviceId", "ClientProcessName",
        "ClientVersion", "SessionId", "ActorInfoString",
    ] if audit.get(k) is not None and str(audit.get(k)).strip() != ""}
    if cl_vals:
        with st.expander("Client & Logon", expanded=False):
            for k, v in cl_vals.items():
                render_kv(k, v)

    # App Access Context
    aac = audit.get("AppAccessContext")
    if aac and isinstance(aac, dict):
        with st.expander("App Access Context", expanded=False):
            for k, v in aac.items():
                render_kv(k, v)

    # Operation Properties
    op_props = audit.get("OperationProperties", [])
    if op_props:
        with st.expander("Operation Properties", expanded=False):
            render_extra_props(op_props)

    # Messages (for MailItemsAccessed)
    messages = audit.get("Messages", [])
    if messages:
        with st.expander(f"Messages ({len(messages)})", expanded=False):
            st.json(messages)

    # SIDs
    sid_fields = [
        ("Logon User SID", "LogonUserSid"),
        ("Mailbox Owner SID", "MailboxOwnerSid"),
        ("Mailbox Owner Master Account SID", "MailboxOwnerMasterAccountSid"),
    ]
    sid_present = [(label, audit.get(key)) for label, key in sid_fields if audit.get(key)]
    if sid_present:
        with st.expander("Security Identifiers", expanded=False):
            for label, val in sid_present:
                render_kv(label, val)

    # Server
    server = audit.get("OriginatingServer")
    if server and not is_cmdlet:
        with st.expander("Server", expanded=False):
            render_kv("Originating Server", server.strip())

# =========================================================================
# SharePoint / OneDrive (same schema)
# =========================================================================
elif workload in ("SharePoint", "OneDrive"):
    # File/Folder info
    with st.expander("File / Folder", expanded=True):
        render_kv("File Name", audit.get("SourceFileName"))
        render_kv("File Extension", audit.get("SourceFileExtension"))
        render_kv("Relative URL", audit.get("SourceRelativeUrl"))
        render_kv("Item Type", audit.get("ItemType"))
        render_kv("File Size", f"{audit.get('FileSizeBytes')} bytes" if audit.get("FileSizeBytes") else None)
        render_kv("List Name", audit.get("ListName") or audit.get("ListTitle"))
        render_kv("List URL", audit.get("ListUrl"))
        # Destination (for moves/renames)
        render_kv("Destination File", audit.get("DestinationFileName"))
        render_kv("Destination Extension", audit.get("DestinationFileExtension"))
        render_kv("Destination URL", audit.get("DestinationRelativeUrl"))

    # Site info
    with st.expander("Site", expanded=False):
        render_kv("Site URL", audit.get("SiteUrl"))
        render_kv("Site Title", audit.get("SiteTitle"))
        render_kv("Site Owner", audit.get("SiteOwnerEmail") or audit.get("SiteOwner"))
        render_kv("Site Template", audit.get("SiteTemplate"))
        render_kv("Site ID", audit.get("Site"))
        render_kv("Web ID", audit.get("WebId"))

    # Sharing details
    sharing_keys = ["TargetUserOrGroupName", "TargetUserOrGroupType", "SharingType",
                    "SharingLinkScope", "Permission", "UniqueSharingId", "EventData"]
    sharing_vals = {k: audit.get(k) for k in sharing_keys if audit.get(k) is not None and str(audit.get(k)).strip() not in ("", "None")}
    if sharing_vals:
        with st.expander("Sharing Details", expanded=True):
            for k, v in sharing_vals.items():
                render_kv(k, v)

    # Search (for SearchQueryPerformed)
    if audit.get("SearchQueryText") or audit.get("QueryText"):
        with st.expander("Search Query", expanded=True):
            render_kv("Query", audit.get("SearchQueryText") or audit.get("QueryText"))
            render_kv("Query Source", audit.get("QuerySource"))

    # Client info
    with st.expander("Client", expanded=False):
        render_kv("Platform", audit.get("Platform"))
        render_kv("Browser", f"{audit.get('BrowserName', '')} {audit.get('BrowserVersion', '')}".strip() or None)
        render_kv("User Agent", audit.get("UserAgent") or audit.get("ClientUserAgent"))
        render_kv("Device", audit.get("DeviceDisplayName"))
        render_kv("Machine ID", audit.get("MachineId"))
        render_kv("Managed Device", audit.get("IsManagedDevice"))
        render_kv("Geo Location", audit.get("GeoLocation"))
        render_kv("Authentication", audit.get("AuthenticationType"))
        render_kv("Application", audit.get("ApplicationDisplayName"))
        render_kv("Application ID", audit.get("ApplicationId"))
        render_kv("Event Source", audit.get("EventSource"))
        render_kv("From App", audit.get("FromApp"))

    # App Access Context
    aac = audit.get("AppAccessContext")
    if aac and isinstance(aac, dict):
        with st.expander("App Access Context", expanded=False):
            for k, v in aac.items():
                render_kv(k, v)

    # Modified Properties (some SP ops have these)
    mod_props = audit.get("ModifiedProperties", [])
    if mod_props:
        with st.expander("Modified Properties", expanded=False):
            render_modified_props(mod_props)

    # Virus info (FileMalwareDetected)
    if audit.get("VirusInfo") or audit.get("VirusVendor"):
        with st.expander("Malware Detection", expanded=True):
            render_kv("Virus Info", audit.get("VirusInfo"))
            render_kv("Virus Vendor", audit.get("VirusVendor"))

# =========================================================================
# Microsoft Teams
# =========================================================================
elif workload == "MicrosoftTeams":
    # Meeting / Call details
    if operation in ("MeetingDetail", "MeetingParticipantDetail", "CallParticipantDetail"):
        with st.expander("Meeting / Call", expanded=True):
            render_kv("Item Name", audit.get("ItemName"))
            render_kv("Meeting URL", audit.get("MeetingURL"))
            render_kv("Conference URI", audit.get("ConferenceUri"))
            render_kv("Call ID", audit.get("CallId"))
            render_kv("Meeting Detail ID", audit.get("MeetingDetailId"))
            render_kv("Communication Type", audit.get("CommunicationType"))
            render_kv("Communication SubType", audit.get("CommunicationSubType"))
            render_kv("Join Time", audit.get("JoinTime"))
            render_kv("Leave Time", audit.get("LeaveTime"))
            render_kv("Start Time", audit.get("StartTime"))
            render_kv("End Time", audit.get("EndTime"))
            render_kv("Device", audit.get("DeviceInformation"))
            render_kv("Modalities", ", ".join(audit.get("Modalities", [])) if audit.get("Modalities") else None)

        # Attendees
        attendees = audit.get("Attendees", [])
        if attendees:
            with st.expander(f"Attendees ({len(attendees)})", expanded=True):
                render_attendees(attendees)

        # Organizer
        organizer = audit.get("Organizer")
        if organizer and isinstance(organizer, dict):
            with st.expander("Organizer", expanded=False):
                for k, v in organizer.items():
                    render_kv(k, v)

        # Participant Info
        pinfo = audit.get("ParticipantInfo")
        if pinfo and isinstance(pinfo, dict):
            with st.expander("Participant Info", expanded=False):
                for k, v in pinfo.items():
                    if isinstance(v, list):
                        render_kv(k, ", ".join(str(x) for x in v) if v else "(none)")
                    else:
                        render_kv(k, v)

    # Message operations
    elif "Message" in operation or operation in ("ChatCreated", "ReactedToMessage"):
        with st.expander("Message", expanded=True):
            render_kv("Chat Name", audit.get("ChatName"))
            render_kv("Chat Thread ID", audit.get("ChatThreadId"))
            render_kv("Message ID", audit.get("MessageId"))
            render_kv("Message Version", audit.get("MessageVersion"))
            render_kv("Communication Type", audit.get("CommunicationType"))
            render_kv("Reaction Type", audit.get("MessageReactionType"))
            render_kv("Is Copilot Mentioned", audit.get("IsCopilotMentioned"))
            # URLs and links
            urls = audit.get("MessageURLs", [])
            if urls:
                st.markdown("**URLs:**")
                for u in urls:
                    st.markdown(f"- {u}")
            links = audit.get("MessageLinks", [])
            if links:
                st.markdown("**Links:**")
                st.json(links)
            files = audit.get("MessageFiles", [])
            if files:
                st.markdown("**Files:**")
                st.json(files)

    # Team / Channel membership
    elif operation in ("MemberAdded", "MemberRemoved", "TeamCreated", "ChannelAdded", "AppInstalled"):
        with st.expander("Team / Channel", expanded=True):
            render_kv("Team Name", audit.get("TeamName"))
            render_kv("Team GUID", audit.get("TeamGuid"))
            render_kv("AAD Group ID", audit.get("AADGroupId"))
            render_kv("Add-On Name", audit.get("AddOnName"))
            render_kv("Add-On Type", audit.get("AddOnType"))
            render_kv("Distribution Mode", audit.get("AppDistributionMode"))
        members = audit.get("Members", [])
        if members:
            with st.expander(f"Members ({len(members)})", expanded=True):
                rows = []
                for m in members:
                    if isinstance(m, dict):
                        rows.append({
                            "Name": m.get("DisplayName", ""),
                            "UPN": m.get("UPN", ""),
                            "Role": m.get("Role", ""),
                        })
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Session events
    elif operation == "TeamsSessionStarted":
        with st.expander("Session", expanded=True):
            render_kv("Device ID", audit.get("DeviceId"))
            render_kv("Target User", audit.get("TargetUserId"))

    else:
        # Generic Teams fields
        with st.expander("Details", expanded=True):
            render_kv("Chat Name", audit.get("ChatName"))
            render_kv("Item Name", audit.get("ItemName"))
            render_kv("Team Name", audit.get("TeamName"))
            render_kv("Communication Type", audit.get("CommunicationType"))
            render_kv("Target User", audit.get("TargetUserId"))

    # Extra Properties (common in Teams)
    extras = audit.get("ExtraProperties", [])
    if extras:
        with st.expander("Extra Properties", expanded=False):
            render_extra_props(extras, use_key_value=True)

    # App Access Context
    aac = audit.get("AppAccessContext")
    if aac and isinstance(aac, dict):
        with st.expander("App Access Context", expanded=False):
            for k, v in aac.items():
                render_kv(k, v)

    # Artifacts Shared
    artifacts = audit.get("ArtifactsShared", [])
    if artifacts:
        with st.expander("Shared Artifacts", expanded=False):
            st.json(artifacts)

# =========================================================================
# Copilot
# =========================================================================
elif workload == "Copilot":
    with st.expander("Copilot Interaction", expanded=True):
        render_kv("App Identity", audit.get("AppIdentity"))
        render_kv("Client Region", audit.get("ClientRegion"))
        render_kv("Log Version", audit.get("CopilotLogVersion"))

    ced = audit.get("CopilotEventData")
    if ced and isinstance(ced, dict):
        with st.expander("Event Data", expanded=True):
            render_kv("App Host", ced.get("AppHost"))
            render_kv("License Type", ced.get("LicenseType"))
            render_kv("Thread ID", ced.get("ThreadId"))

            # AI Plugins
            plugins = ced.get("AISystemPlugin", [])
            if plugins:
                st.markdown("**AI Plugins:**")
                for p in plugins:
                    st.markdown(f"- {p.get('Id', '')} ({p.get('Name', '')})")

            # Messages
            msgs = ced.get("Messages", [])
            if msgs:
                st.markdown("**Messages:**")
                rows = []
                for m in msgs:
                    rows.append({
                        "ID": m.get("Id", ""),
                        "Is Prompt": m.get("isPrompt", ""),
                        "Jailbreak Detected": m.get("JailbreakDetected", ""),
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # Model details
            models = ced.get("ModelTransparencyDetails", [])
            if models:
                st.markdown("**Models:** " + ", ".join(m.get("ModelName", "") for m in models))

            # Accessed Resources
            resources = ced.get("AccessedResources", [])
            if resources:
                with st.expander("Accessed Resources", expanded=False):
                    st.json(resources)

            # Contexts
            contexts = ced.get("Contexts", [])
            if contexts:
                with st.expander("Contexts", expanded=False):
                    st.json(contexts)

# =========================================================================
# Security & Compliance Center
# =========================================================================
elif workload == "SecurityComplianceCenter":
    # Audit search operations
    if "AuditSearch" in operation:
        with st.expander("Audit Search", expanded=True):
            render_kv("Search Job Name", audit.get("SearchJobName"))
            render_kv("Search Job ID", audit.get("SearchJobId"))
            render_kv("Search Source", audit.get("SearchSource"))
            render_kv("Completion Status", audit.get("CompletionStatus"))
            render_kv("Results Count", audit.get("ResultsCount") or audit.get("ResultCount"))
            render_kv("Start Time", audit.get("StartTime"))
            # Parse SearchFilters if present
            sf_raw = audit.get("SearchFilters")
            if sf_raw:
                sf = try_parse_json(sf_raw)
                if isinstance(sf, dict):
                    st.markdown("**Search Filters:**")
                    for k, v in sf.items():
                        if isinstance(v, list) and len(v) > 10:
                            render_kv(k, f"({len(v)} items)")
                        else:
                            render_kv(k, v)
                else:
                    st.text(str(sf_raw))

    # Alert
    elif operation == "AlertTriggered":
        with st.expander("Alert", expanded=True):
            render_kv("Alert ID", audit.get("AlertId"))
            render_kv("Alert Type", audit.get("AlertType"))
            render_kv("Category", audit.get("Category"))
            render_kv("Severity", audit.get("Severity"))
            render_kv("Status", audit.get("Status"))
            render_kv("Comments", audit.get("Comments"))
            links = audit.get("AlertLinks", [])
            if links:
                st.markdown("**Links:**")
                st.json(links)

    # User Submission (spam/phishing reports)
    elif operation in ("UserSubmission", "UserSubmissionTriage"):
        with st.expander("Submission", expanded=True):
            render_kv("Subject", audit.get("Subject"))
            render_kv("Internet Message ID", audit.get("InternetMessageId"))
            render_kv("P1 Sender", audit.get("P1Sender"))
            render_kv("P2 Sender", audit.get("P2Sender"))
            render_kv("P1 Sender Domain", audit.get("P1SenderDomain"))
            render_kv("P2 Sender Domain", audit.get("P2SenderDomain"))
            render_kv("Sender IP", audit.get("SenderIP"))
            render_kv("Language", audit.get("Language"))
            render_kv("BCL Value", audit.get("BCLValue"))
            render_kv("Message Date", audit.get("MessageDate"))
            render_kv("Filtering Date", audit.get("FilteringDate"))
        # Delivery info
        dmi = audit.get("DeliveryMessageInfo")
        if dmi and isinstance(dmi, dict):
            with st.expander("Delivery Info", expanded=True):
                for k, v in dmi.items():
                    render_kv(k, v)

    else:
        # Generic SCC fields
        with st.expander("Details", expanded=True):
            render_kv("Name", audit.get("Name"))
            render_kv("Data Type", audit.get("DataType"))
            render_kv("Client Application", audit.get("ClientApplication"))
            render_kv("Status", audit.get("Status"))

    # SCC Parameters
    params = audit.get("Parameters") or audit.get("NonPIIParameters")
    if params and isinstance(params, list):
        with st.expander("Parameters", expanded=False):
            render_parameters(params)

    # Extended Properties
    ext_props = audit.get("ExtendedProperties", [])
    if ext_props:
        with st.expander("Extended Properties", expanded=False):
            render_extended_props(ext_props)

# =========================================================================
# Quarantine
# =========================================================================
elif workload == "Quarantine":
    with st.expander("Quarantine Details", expanded=True):
        render_kv("Network Message ID", audit.get("NetworkMessageId"))
        render_kv("Release To", audit.get("ReleaseTo"))
        render_kv("Request Type", audit.get("RequestType"))
        render_kv("Request Source", audit.get("RequestSource"))

# =========================================================================
# Microsoft To Do
# =========================================================================
elif workload == "MicrosoftTodo":
    with st.expander("Task Details", expanded=True):
        render_kv("Item Type", audit.get("ItemType"))
        render_kv("Item ID", audit.get("ItemId"))
        render_kv("Actor App ID", audit.get("ActorAppId"))
        render_kv("Target Actor", audit.get("TargetActorId"))
        render_kv("Target Tenant", audit.get("TargetActorTenantId"))

    extras = audit.get("ExtraProperties", [])
    if extras:
        with st.expander("Extra Properties", expanded=False):
            render_extra_props(extras, use_key_value=True)

# =========================================================================
# Planner
# =========================================================================
elif workload == "Planner":
    with st.expander("Planner Details", expanded=True):
        render_kv("Plan ID", audit.get("PlanId"))
        render_kv("Plan List", audit.get("PlanList"))
        render_kv("Object ID", audit.get("ObjectId"))

    aac = audit.get("AppAccessContext")
    if aac and isinstance(aac, dict):
        with st.expander("App Access Context", expanded=False):
            for k, v in aac.items():
                render_kv(k, v)

# =========================================================================
# Yammer / Viva Engage
# =========================================================================
elif workload == "Yammer":
    with st.expander("Yammer Details", expanded=True):
        render_kv("Actor User ID", audit.get("ActorUserId"))
        render_kv("Actor Yammer User ID", audit.get("ActorYammerUserId"))
        render_kv("Yammer Network ID", audit.get("YammerNetworkId"))
        render_kv("Details", audit.get("Details"))

# =========================================================================
# Generic fallback for unknown workloads
# =========================================================================
else:
    # Actor
    actors = audit.get("Actor", [])
    if actors:
        with st.expander("Actor", expanded=True):
            st.dataframe(format_actor_target(actors), use_container_width=True, hide_index=True)

    # Target
    targets = audit.get("Target", [])
    if targets:
        with st.expander("Target", expanded=True):
            st.dataframe(format_actor_target(targets), use_container_width=True, hide_index=True)

    # Modified Properties
    mod_props = audit.get("ModifiedProperties", [])
    if mod_props:
        with st.expander("Modified Properties", expanded=True):
            render_modified_props(mod_props)

    # Extended Properties
    ext_props = audit.get("ExtendedProperties", [])
    if ext_props:
        with st.expander("Extended Properties", expanded=False):
            render_extended_props(ext_props)

    # Item (generic)
    item = audit.get("Item")
    if item:
        with st.expander("Item Details", expanded=False):
            st.json(item)

# Full JSON (always available)
with st.expander("Full AuditData JSON", expanded=False):
    st.json(audit)
