import io
import re
from urllib.parse import urlparse
import openpyxl
import pandas as pd
import streamlit as st

# ==============================================================================
# PAGE CONFIG & STYLING
# ==============================================================================
st.set_page_config(
    page_title="ISA Auto-Compiler v2.0 | Technical Audit Engine",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Initial Site Analysis (ISA) Auto-Compiler v2.0")
st.markdown(
    "Upload your **Initial Site Analysis Spreadsheet Template (.xlsx)** alongside your "
    "Screaming Frog exports (`internal_all`, `all_inlinks`, `all_redirects`, `images_all`) to "
    "automatically populate 35+ evidence tabs and flag the master checklist."
)

# ==============================================================================
# DEFENSIVE FILE PARSERS & HELPERS
# ==============================================================================
def load_csv_robust(uploaded_file) -> pd.DataFrame:
    """Safely parses Screaming Frog / GSC / Backlink CSVs across UTF-8, BOM, and UTF-16."""
    if uploaded_file is None:
        return pd.DataFrame()
    raw_bytes = uploaded_file.getvalue()

    # Screaming Frog Redirect Chains exports sometimes include a metadata title row on Row 1
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "latin1"):
        for sep in (",", "\t", ";"):
            for skip in (0, 1):
                try:
                    df = pd.read_csv(
                        io.BytesIO(raw_bytes),
                        encoding=encoding,
                        sep=sep,
                        skiprows=skip,
                        low_memory=False,
                    )
                    if len(df.columns) > 1 and not all(str(c).startswith("Unnamed") for c in df.columns):
                        df.columns = [str(c).strip() for c in df.columns]
                        return df
                except Exception:
                    continue
    return pd.DataFrame()


def find_col(df: pd.DataFrame, exact_candidates: list[str], allow_substring: bool = False) -> str | None:
    """Strict case-insensitive column resolver to avoid false substring matches."""
    if df.empty:
        return None
    lower_map = {str(c).lower().strip(): c for c in df.columns}
    for cand in exact_candidates:
        c_low = cand.lower().strip()
        if c_low in lower_map:
            return lower_map[c_low]
    if allow_substring:
        for cand in exact_candidates:
            c_low = cand.lower().strip()
            for col_low, orig in lower_map.items():
                if c_low in col_low:
                    return orig
    return None


def clean_str_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def is_html_page(df: pd.DataFrame) -> pd.Series:
    ct_col = find_col(df, ["Content Type", "Content"])
    if not ct_col:
        return pd.Series(True, index=df.index)
    return clean_str_series(df[ct_col]).str.lower().str.contains("text/html", na=False)


def suggest_parent_redirect(url: str, live_urls_set: set[str]) -> str:
    """Walks up URL directory segments to suggest the closest live 200 OK parent URL."""
    try:
        parsed = urlparse(url)
        base = f"{parsed.scheme}://{parsed.netloc}"
        parts = [p for p in parsed.path.split("/") if p]
        while len(parts) > 1:
            parts.pop()
            cand_slash = f"{base}/{'/'.join(parts)}/"
            cand_noslash = f"{base}/{'/'.join(parts)}"
            if cand_slash in live_urls_set:
                return cand_slash
            if cand_noslash in live_urls_set:
                return cand_noslash
        if f"{base}/" in live_urls_set:
            return f"{base}/"
        if base in live_urls_set:
            return base
    except Exception:
        pass
    return ""


def suggest_clean_slug(url: str) -> str:
    """Generates a lowercase, hyphenated clean URL suggestion."""
    try:
        parsed = urlparse(url)
        clean_path = re.sub(r"[^a-z0-9/._-]+", "-", parsed.path.lower())
        clean_path = re.sub(r"-{2,}", "-", clean_path)
        return f"{parsed.scheme.lower()}://{parsed.netloc.lower()}{clean_path}"
    except Exception:
        return url.lower()


def suggest_alt_from_filename(img_url: str) -> str:
    """Extracts human-readable alt text from an image filename."""
    try:
        path = urlparse(img_url).path
        fname = path.split("/")[-1]
        stem = re.sub(r"\.[a-zA-Z0-9]+$", "", fname)
        words = re.sub(r"[-_]+", " ", stem).strip()
        if not words or re.match(r"^(img|dsc|screenshot|image|photo|untitled|\d+)$", words, re.I):
            return "[Add descriptive product/context alt text]"
        return words.capitalize()
    except Exception:
        return "[Add descriptive alt text]"


# ==============================================================================
# CORE AUDIT ENGINE
# ==============================================================================
def run_audit_engine(
    df_internal: pd.DataFrame,
    df_inlinks: pd.DataFrame,
    df_redirects: pd.DataFrame,
    df_images: pd.DataFrame,
    df_gsc: pd.DataFrame,
    df_backlinks: pd.DataFrame,
) -> tuple[dict[str, list[list]], dict[str, tuple[str, str]]]:
    """
    Computes all evidence tab rows and master sheet updates keyed by ITEM substring.
    """
    tab_data: dict[str, list[list]] = {}
    master_updates: dict[str, tuple[str, str]] = {}

    ISSUE_YES = "✘Issues found"
    ISSUE_NO = "✔No issues"

    def record_check(
        tab_name: str | None,
        rows: list[list],
        item_keys: list[str],
        issue_msg_template: str,
        clean_msg: str = "No issues detected in crawl export.",
    ):
        if tab_name:
            tab_data[tab_name] = rows
        count = len(rows)
        for key in item_keys:
            if count > 0:
                master_updates[key] = (ISSUE_YES, issue_msg_template.format(count=count))
            else:
                master_updates[key] = (ISSUE_NO, clean_msg)

    # Lookup dictionaries populated from internal_all.csv
    live_200_urls: set[str] = set()
    url_to_canonical: dict[str, str] = {}
    url_to_status: dict[str, int] = {}
    url_to_indexability: dict[str, str] = {}
    url_to_size: dict[str, int] = {}

    # --------------------------------------------------------------------------
    # 1. PARSE INTERNAL_ALL.CSV
    # --------------------------------------------------------------------------
    if not df_internal.empty:
        addr_col = find_col(df_internal, ["Address", "URL"])
        status_col = find_col(df_internal, ["Status Code"])
        index_col = find_col(df_internal, ["Indexability"])
        canon_col = find_col(df_internal, ["Canonical Link Element 1", "Canonical"])
        sitemap_col = find_col(df_internal, ["In Sitemap", "Sitemap"], allow_substring=True)
        inlinks_col = find_col(df_internal, ["Unique Inlinks", "Inlinks"])
        size_col = find_col(df_internal, ["Size (Bytes)", "Size"])

        title_1_col = find_col(df_internal, ["Title 1"])
        title_1_len_col = find_col(df_internal, ["Title 1 Length"])
        title_2_col = find_col(df_internal, ["Title 2"])

        meta_1_col = find_col(df_internal, ["Meta Description 1"])
        meta_1_len_col = find_col(df_internal, ["Meta Description 1 Length"])
        meta_2_col = find_col(df_internal, ["Meta Description 2"])
        meta_kw_col = find_col(df_internal, ["Meta Keywords 1", "Meta Keyword 1"])

        h1_1_col = find_col(df_internal, ["H1-1", "H1 1"])
        h1_2_col = find_col(df_internal, ["H1-2", "H1 2"])
        rel_next_col = find_col(df_internal, ["Rel Next 1", " rel=\"next\" 1"])
        rel_prev_col = find_col(df_internal, ["Rel Prev 1", " rel=\"prev\" 1"])
        og_title_col = find_col(df_internal, ["Open Graph Title", "og:title", "og:title 1"], allow_substring=True)
        og_url_col = find_col(df_internal, ["Open Graph URL", "og:url", "og:url 1"], allow_substring=True)

        df_int = df_internal.copy()
        df_int["_url"] = clean_str_series(df_int[addr_col]) if addr_col else ""
        df_int["_status"] = pd.to_numeric(df_int[status_col], errors="coerce").fillna(0).astype(int) if status_col else 0
        df_int["_indexability"] = clean_str_series(df_int[index_col]) if index_col else ""
        df_int["_canonical"] = clean_str_series(df_int[canon_col]) if canon_col else ""
        df_int["_size"] = pd.to_numeric(df_int[size_col], errors="coerce").fillna(0).astype(int) if size_col else 0
        df_int["_is_html"] = is_html_page(df_int)

        for _, row in df_int.iterrows():
            u = row["_url"]
            st_code = row["_status"]
            url_to_status[u] = st_code
            url_to_canonical[u] = row["_canonical"]
            url_to_indexability[u] = row["_indexability"]
            url_to_size[u] = row["_size"]
            if st_code == 200:
                live_200_urls.add(u)

        # Internal Pages | 5xx Errors
        rows_5xx = [[u] for u in df_int.loc[df_int["_status"].between(500, 599), "_url"].tolist()]
        record_check(
            "Pages returning 5xx Errors",
            rows_5xx,
            ["Pages Returning 5xx Errors"],
            "Found {count} internal URL(s) returning 5xx server errors (see 'Pages returning 5xx Errors' tab).",
        )

        # Internal Pages | 4xx Errors
        urls_4xx = df_int.loc[df_int["_status"].between(400, 499), "_url"].tolist()
        rows_4xx = [[u, suggest_parent_redirect(u, live_200_urls)] for u in urls_4xx]
        record_check(
            "Pages returning 4xx Errors",
            rows_4xx,
            ["Pages Returning 4xx Errors"],
            "Found {count} internal URL(s) returning 4xx errors (see 'Pages returning 4xx Errors' tab).",
        )

        # HTTPS checks
        http_urls = df_int.loc[df_int["_url"].str.startswith("http://"), "_url"].tolist()
        if http_urls:
            master_updates["Website is not using HTTPS"] = (ISSUE_YES, f"Found {len(http_urls)} internal URL(s) served over insecure HTTP.")
            master_updates["HTTP vs. HTTPS"] = (ISSUE_YES, f"Found {len(http_urls)} HTTP URL(s) accessible alongside HTTPS.")
        else:
            master_updates["Website is not using HTTPS"] = (ISSUE_NO, "All crawled internal URLs use HTTPS.")
            master_updates["HTTP vs. HTTPS"] = (ISSUE_NO, "No duplicate HTTP URLs detected.")

        # www vs non-www check
        hosts = {urlparse(u).netloc.lower() for u in df_int["_url"] if u.startswith("http")}
        has_www = any(h.startswith("www.") for h in hosts)
        has_non_www = any(not h.startswith("www.") for h in hosts)
        if has_www and has_non_www:
            master_updates["Duplicates due to www and non-www"] = (ISSUE_YES, f"Both www and non-www hosts crawled: {', '.join(sorted(hosts))}.")
        else:
            master_updates["Duplicates due to www and non-www"] = (ISSUE_NO, f"Consistent host version detected ({', '.join(sorted(hosts))}).")

        df_html = df_int[df_int["_is_html"]].copy()
        df_html_200 = df_html[df_html["_status"] == 200].copy()

        # XML Sitemap Checks
        if sitemap_col:
            in_sm = clean_str_series(df_int[sitemap_col]).str.lower().isin(["true", "yes", "1"])
            df_sm = df_int[in_sm]

            sm_4xx = [[u] for u in df_sm.loc[df_sm["_status"].between(400, 499), "_url"].tolist()]
            record_check(
                "XML Sitemap - 4xx error ",
                sm_4xx,
                ["XML Sitemap  | 4xx errors"],
                "Found {count} 4xx URL(s) inside XML Sitemap (see 'XML Sitemap - 4xx error ' tab).",
            )

            sm_5xx = [[u] for u in df_sm.loc[df_sm["_status"].between(500, 599), "_url"].tolist()]
            record_check(
                "XML Sitemap - 5xx error",
                sm_5xx,
                ["XML Sitemap  | 5xx errors"],
                "Found {count} 5xx URL(s) inside XML Sitemap (see 'XML Sitemap - 5xx error' tab).",
            )

            sm_non_canon = []
            for _, r in df_sm[df_sm["_status"] == 200].iterrows():
                u, c = r["_url"], r["_canonical"]
                if (c and c != u) or ("non-indexable" in r["_indexability"].lower()):
                    sm_non_canon.append([u])
            record_check(
                "XML Sitemap - Non canonical URl",
                sm_non_canon,
                ["XML Sitemap  | Non-canonical URLs"],
                "Found {count} non-canonical/non-indexable URL(s) in XML Sitemap (see 'XML Sitemap - Non canonical URl' tab).",
            )

            sm_missing = []
            for _, r in df_html_200[~in_sm.loc[df_html_200.index]].iterrows():
                u, c = r["_url"], r["_canonical"]
                if r["_indexability"].lower() == "indexable" and (not c or c == u):
                    sm_missing.append([u])
            record_check(
                "XML Sitemap - Important pages m",
                sm_missing,
                ["XML Sitemap  | Important pages missing"],
                "Found {count} indexable 200 OK HTML page(s) missing from XML Sitemap (see 'XML Sitemap - Important pages m' tab).",
            )

        # Orphan Pages
        if inlinks_col:
            inlink_counts = pd.to_numeric(df_html[inlinks_col], errors="coerce").fillna(0)
            orphans = df_html.loc[inlink_counts == 0, "_url"].tolist()
            if orphans:
                master_updates["Internal Pages | Orphan Pages"] = (ISSUE_YES, f"Found {len(orphans)} HTML page(s) with 0 internal inlinks.")
            else:
                master_updates["Internal Pages | Orphan Pages"] = (ISSUE_NO, "0 orphan HTML pages detected in crawl.")

        # Ugly URLs & Uppercase in URLs
        ugly_rows, upper_rows = [], []
        for u in df_html["_url"].unique():
            parsed = urlparse(u)
            path_and_q = f"{parsed.path}?{parsed.query}" if parsed.query else parsed.path
            if any(c.isupper() for c in path_and_q):
                upper_rows.append([u, suggest_clean_slug(u)])
            if any(ch in path_and_q for ch in ["?", "&", "=", "_"]) or re.search(r"/\d{4,}/?$", parsed.path):
                ugly_rows.append([u, suggest_clean_slug(u)])
        tab_data["Uppercase in URLs"] = upper_rows
        tab_data["Ugly URLs"] = ugly_rows

        # Duplicates from Non-canonicalization
        missing_canon_rows = [[r["_url"], r["_url"]] for _, r in df_html_200.iterrows() if not r["_canonical"]]
        record_check(
            "Duplicates from Non-canonicaliz",
            missing_canon_rows,
            ["Canonical Errors"],
            "Found {count} 200 OK HTML page(s) missing a canonical tag (see 'Duplicates from Non-canonicaliz' tab).",
        )

        # Duplicates from trailing slashes
        all_html_set = set(df_html["_url"].tolist())
        trailing_rows = []
        seen_pairs = set()
        for u in sorted(all_html_set):
            parsed = urlparse(u)
            if parsed.path in ("", "/"):
                continue
            alt = u.rstrip("/") if u.endswith("/") else u + "/"
            if alt in all_html_set:
                pair_key = tuple(sorted([u, alt]))
                if pair_key not in seen_pairs:
                    seen_pairs.add(pair_key)
                    canon_suggest = url_to_canonical.get(u) or (u if u.endswith("/") else alt)
                    trailing_rows.append([u, canon_suggest])
                    trailing_rows.append([alt, canon_suggest])
        record_check(
            "Duplicates from trailing slashe",
            trailing_rows,
            ["Duplicates from trailing slashes"],
            "Found {count} URL(s) accessible with both trailing and non-trailing slashes (see 'Duplicates from trailing slashe' tab).",
        )

        # Pagination
        pag_rows = []
        pag_pattern = re.compile(r"(/page/\d+|[?&](page|p|paged)=\d+)", re.I)
        rel_next_s = clean_str_series(df_html_200[rel_next_col]) if rel_next_col else pd.Series("", index=df_html_200.index)
        rel_prev_s = clean_str_series(df_html_200[rel_prev_col]) if rel_prev_col else pd.Series("", index=df_html_200.index)
        for idx, r in df_html_200.iterrows():
            u = r["_url"]
            if pag_pattern.search(u) and not rel_next_s.loc[idx] and not rel_prev_s.loc[idx]:
                pag_rows.append([u, r["_canonical"] or u])
        record_check(
            "Pagination",
            pag_rows,
            ["Pagination"],
            "Found {count} paginated URL(s) missing rel='prev'/rel='next' attributes (see 'Pagination' tab).",
        )

        # H1 Tags & H1 same as Title Tag
        h1_rows = []
        h1_same_title_count = 0
        if h1_1_col:
            h1_1_s = clean_str_series(df_html_200[h1_1_col])
            h1_2_s = clean_str_series(df_html_200[h1_2_col]) if h1_2_col else pd.Series("", index=df_html_200.index)
            title_s_for_h1 = clean_str_series(df_html_200[title_1_col]) if title_1_col else pd.Series("", index=df_html_200.index)

            h1_counts = h1_1_s[h1_1_s != ""].value_counts()
            dup_h1s = set(h1_counts[h1_counts > 1].index)

            for idx, r in df_html_200.iterrows():
                u = r["_url"]
                h1_val = h1_1_s.loc[idx]
                h2_val = h1_2_s.loc[idx]
                t_val = title_s_for_h1.loc[idx]

                if h1_val and t_val and h1_val.lower() == t_val.lower():
                    h1_same_title_count += 1

                issues = []
                if not h1_val:
                    issues.append("Missing H1")
                if h2_val:
                    issues.append("Multiple H1 Tags")
                if h1_val in dup_h1s:
                    issues.append("Duplicate H1")
                if issues:
                    h1_rows.append([u, ", ".join(issues), "Ensure a single, unique, keyword-targeted H1 tag."])

        record_check(
            "H1 Tags",
            h1_rows,
            ["Headers & H1 Tags"],
            "Found {count} page(s) with Missing, Multiple, or Duplicate H1 tags (see 'H1 Tags' tab).",
        )
        master_updates["H1 same as Title Tag"] = (
            (ISSUE_YES, f"Found {h1_same_title_count} page(s) where H1 is identical to the Title tag.")
            if h1_same_title_count > 0
            else (ISSUE_NO, "H1 tags are differentiated from Title tags.")
        )

        # Title Tags: Missing, Duplicate, Multiple (Uses Title 2 to detect Multiple!)
        t_missing_rows, t_dup_rows, t_mult_rows = [], [], []
        if title_1_col:
            t1_s = clean_str_series(df_html_200[title_1_col])
            t2_s = clean_str_series(df_html_200[title_2_col]) if title_2_col else pd.Series("", index=df_html_200.index)
            t1_len_s = (
                pd.to_numeric(df_html_200[title_1_len_col], errors="coerce").fillna(t1_s.str.len()).astype(int)
                if title_1_len_col
                else t1_s.str.len()
            )
            t_counts = t1_s[t1_s != ""].value_counts()
            dup_titles = set(t_counts[t_counts > 1].index)

            for idx, r in df_html_200.iterrows():
                u = r["_url"]
                tv = t1_s.loc[idx]
                tv2 = t2_s.loc[idx]
                tl = int(t1_len_s.loc[idx])
                if not tv:
                    t_missing_rows.append([u, "", 0, "Missing Title Tag"])
                else:
                    if tv in dup_titles:
                        t_dup_rows.append([u, tv, tl, "Duplicate Title Tag"])
                    if tv2:
                        t_mult_rows.append([u, tv, tl, f"Multiple Title Tags (Title 2: {tv2[:50]})"])

        tab_data["Title tags - Missing"] = t_missing_rows
        tab_data["Title tags - Duplicate"] = t_dup_rows
        tab_data["Title tags - Multiple"] = t_mult_rows
        total_title_issues = len(t_missing_rows) + len(t_dup_rows) + len(t_mult_rows)
        master_updates["Title Tags (duplicates / multiple / absent)"] = (
            (ISSUE_YES, f"Missing: {len(t_missing_rows)} | Duplicate: {len(t_dup_rows)} | Multiple: {len(t_mult_rows)} (see Title tag tabs).")
            if total_title_issues > 0
            else (ISSUE_NO, "No missing, duplicate, or multiple Title tags found.")
        )

        # Meta Descriptions: Too Long, Too Short, Missing, Duplicate, Multiple (Uses Meta Description 2!)
        m_long_rows, m_short_rows, m_miss_rows, m_dup_rows, m_mult_rows = [], [], [], [], []
        if meta_1_col:
            m1_s = clean_str_series(df_html_200[meta_1_col])
            m2_s = clean_str_series(df_html_200[meta_2_col]) if meta_2_col else pd.Series("", index=df_html_200.index)
            m1_len_s = (
                pd.to_numeric(df_html_200[meta_1_len_col], errors="coerce").fillna(m1_s.str.len()).astype(int)
                if meta_1_len_col
                else m1_s.str.len()
            )
            m_counts = m1_s[m1_s != ""].value_counts()
            dup_metas = set(m_counts[m_counts > 1].index)

            for idx, r in df_html_200.iterrows():
                u = r["_url"]
                mv = m1_s.loc[idx]
                mv2 = m2_s.loc[idx]
                ml = int(m1_len_s.loc[idx])
                if not mv:
                    m_miss_rows.append([u, "", 0, "Missing Meta Description"])
                else:
                    if ml > 160:
                        m_long_rows.append([u, mv, ml, f"Over 160 characters ({ml} chars)"])
                    elif ml < 70:
                        m_short_rows.append([u, mv, ml, f"Under 70 characters ({ml} chars)"])
                    if mv in dup_metas:
                        m_dup_rows.append([u, mv, ml, "Duplicate Meta Description"])
                    if mv2:
                        m_mult_rows.append([u, mv, ml, "Multiple Meta Descriptions on page"])

        tab_data["Meta descriptions - Too Long"] = m_long_rows
        tab_data["Meta descriptions - Too Short"] = m_short_rows
        tab_data["Meta descriptions - Missing"] = m_miss_rows
        tab_data["Meta descriptions - Duplicate"] = m_dup_rows
        tab_data["Meta descriptions - Multiple"] = m_mult_rows
        total_meta_issues = len(m_long_rows) + len(m_short_rows) + len(m_miss_rows) + len(m_dup_rows) + len(m_mult_rows)
        master_updates["Meta Descriptions (duplicates / short / long / absent)"] = (
            (ISSUE_YES, f"Missing: {len(m_miss_rows)} | Duplicate: {len(m_dup_rows)} | Long: {len(m_long_rows)} | Short: {len(m_short_rows)} | Multiple: {len(m_mult_rows)}.")
            if total_meta_issues > 0
            else (ISSUE_NO, "No Meta Description issues detected.")
        )

        # Meta Keyword Tags check
        if meta_kw_col:
            mkw_s = clean_str_series(df_html_200[meta_kw_col])
            mkw_count = int((mkw_s != "").sum())
            master_updates["Meta Keyword Tags"] = (
                (ISSUE_YES, f"Found {mkw_count} page(s) still using legacy Meta Keywords tags.")
                if mkw_count > 0
                else (ISSUE_NO, "No Meta Keywords tags detected.")
            )

        # Open Graph Checks
        if og_title_col or og_url_col:
            og_t_s = clean_str_series(df_html_200[og_title_col]) if og_title_col else pd.Series("", index=df_html_200.index)
            og_u_s = clean_str_series(df_html_200[og_url_col]) if og_url_col else pd.Series("", index=df_html_200.index)

            og_missing = int(((og_t_s == "") & (og_u_s == "")).sum())
            og_incomplete = int(((og_t_s == "") ^ (og_u_s == "")).sum())
            og_mismatch = int(((og_u_s != "") & (df_html_200["_canonical"] != "") & (og_u_s != df_html_200["_canonical"])).sum())

            master_updates["Open Graph tags missing"] = (ISSUE_YES, f"Found {og_missing} 200 OK page(s) missing Open Graph tags.") if og_missing else (ISSUE_NO, "Open Graph tags present.")
            master_updates["Open Graph tags incomplete"] = (ISSUE_YES, f"Found {og_incomplete} 200 OK page(s) with incomplete Open Graph tags.") if og_incomplete else (ISSUE_NO, "No incomplete Open Graph tags found.")
            master_updates["Open Graph URL not matching canonical"] = (ISSUE_YES, f"Found {og_mismatch} page(s) where og:url does not match canonical URL.") if og_mismatch else (ISSUE_NO, "og:url tags match canonical URLs.")

    # --------------------------------------------------------------------------
    # 2. PARSE ALL_REDIRECTS.CSV
    # --------------------------------------------------------------------------
    redirect_map: dict[str, tuple[str, int, int]] = {}
    if not df_redirects.empty:
        r_addr_col = find_col(df_redirects, ["Address", "Redirecting URL", "Source", "URL"])
        r_final_col = find_col(df_redirects, ["Final Address", "Final URL", "Destination", "Redirect URL 1"])
        r_fstat_col = find_col(df_redirects, ["Final Status Code", "Status Code"])
        r_hops_col = find_col(df_redirects, ["Number of Redirects", "Redirects", "Hops"])
        r_loop_col = find_col(df_redirects, ["Redirect Loop", "Loop"])

        df_red = df_redirects.copy()
        df_red["_src"] = clean_str_series(df_red[r_addr_col]) if r_addr_col else ""
        df_red["_fin"] = clean_str_series(df_red[r_final_col]) if r_final_col else ""
        df_red["_fstat"] = pd.to_numeric(df_red[r_fstat_col], errors="coerce").fillna(0).astype(int) if r_fstat_col else 0
        df_red["_hops"] = pd.to_numeric(df_red[r_hops_col], errors="coerce").fillna(1).astype(int) if r_hops_col else 1
        df_red["_loop"] = clean_str_series(df_red[r_loop_col]).str.lower().isin(["true", "yes", "1"]) if r_loop_col else False

        for _, r in df_red.iterrows():
            if r["_src"]:
                redirect_map[r["_src"]] = (r["_fin"], r["_fstat"], r["_hops"])

        # Redirects to missing pages
        red_missing = []
        for _, r in df_red[(df_red["_fstat"] >= 400) | (df_red["_fstat"] == 0)].iterrows():
            if r["_src"]:
                red_missing.append([r["_src"], r["_fin"], suggest_parent_redirect(r["_src"], live_200_urls)])
        record_check(
            "Redirects to missing pages",
            red_missing,
            ["Redirects to Missing Pages"],
            "Found {count} redirect(s) resolving to missing/error pages (see 'Redirects to missing pages' tab).",
        )

        # Chains Redirects -> Include full Screaming Frog Redirect Chains headers + rows
        df_chains = df_red[(df_red["_hops"] >= 2) | (df_red["_loop"])].drop(
            columns=["_src", "_fin", "_fstat", "_hops", "_loop"], errors="ignore"
        )
        if not df_chains.empty:
            chain_rows = [df_chains.columns.tolist()] + df_chains.fillna("").values.tolist()
            tab_data["Chains Redirects"] = chain_rows
            master_updates["Internal Pages | Chains Redirects"] = (
                ISSUE_YES,
                f"Found {len(df_chains)} redirect chain(s)/loop(s) with 2+ hops (see 'Chains Redirects' tab).",
            )
        else:
            tab_data["Chains Redirects"] = []
            master_updates["Internal Pages | Chains Redirects"] = (ISSUE_NO, "0 multi-hop redirect chains found.")

        # XML Sitemap - Redirecting
        if not df_internal.empty and find_col(df_internal, ["In Sitemap", "Sitemap"], allow_substring=True):
            sm_col = find_col(df_internal, ["In Sitemap", "Sitemap"], allow_substring=True)
            in_sm_mask = clean_str_series(df_internal[sm_col]).str.lower().isin(["true", "yes", "1"])
            addr_c = find_col(df_internal, ["Address", "URL"])
            sm_urls = clean_str_series(df_internal.loc[in_sm_mask, addr_c]).tolist()
            sm_red_rows = []
            for u in sm_urls:
                if u in redirect_map:
                    sm_red_rows.append([u, redirect_map[u][0]])
                elif 300 <= url_to_status.get(u, 0) < 400:
                    sm_red_rows.append([u, ""])
            record_check(
                "XML Sitemap - Redirecting",
                sm_red_rows,
                ["XML Sitemap  | Redirecting"],
                "Found {count} redirecting URL(s) in the XML Sitemap (see 'XML Sitemap - Redirecting' tab).",
            )

    # --------------------------------------------------------------------------
    # 3. PARSE ALL_INLINKS.CSV (HYPERLINKS + IMAGE INLINKS!)
    # --------------------------------------------------------------------------
    df_img_inlinks = pd.DataFrame()
    if not df_inlinks.empty:
        type_col = find_col(df_inlinks, ["Type"])
        src_col = find_col(df_inlinks, ["Source"])
        dst_col = find_col(df_inlinks, ["Destination"])
        anc_col = find_col(df_inlinks, ["Anchor", "Anchor Text"])
        alt_col = find_col(df_inlinks, ["Alt Text"])
        sc_col = find_col(df_inlinks, ["Status Code"])
        fol_col = find_col(df_inlinks, ["Follow"])
        rel_col = find_col(df_inlinks, ["Rel"])
        pos_col = find_col(df_inlinks, ["Link Position"])

        df_inl_all = df_inlinks.copy()
        df_inl_all["_type"] = clean_str_series(df_inl_all[type_col]).str.lower() if type_col else "hyperlink"
        df_inl_all["_src"] = clean_str_series(df_inl_all[src_col]) if src_col else ""
        df_inl_all["_dst"] = clean_str_series(df_inl_all[dst_col]) if dst_col else ""
        df_inl_all["_anc"] = clean_str_series(df_inl_all[anc_col]) if anc_col else ""
        df_inl_all["_alt"] = clean_str_series(df_inl_all[alt_col]) if alt_col else ""
        df_inl_all["_sc"] = pd.to_numeric(df_inl_all[sc_col], errors="coerce").fillna(0).astype(int) if sc_col else 0
        df_inl_all["_fol"] = clean_str_series(df_inl_all[fol_col]).str.lower() if fol_col else "true"
        df_inl_all["_rel"] = clean_str_series(df_inl_all[rel_col]).str.lower() if rel_col else ""
        df_inl_all["_pos"] = clean_str_series(df_inl_all[pos_col]).str.lower() if pos_col else ""

        # Separate Image Inlinks for the Image Audit Engine
        df_img_inlinks = df_inl_all[df_inl_all["_type"].str.contains("image", na=False)].copy()

        # Filter to Hyperlinks for Link Audits
        df_inl = df_inl_all[df_inl_all["_type"].str.contains("hyperlink|ahref|^a$|^$", regex=True, na=False)].copy()

        site_hosts = {urlparse(u).netloc.lower() for u in df_inl["_src"] if u.startswith("http")}
        is_internal_dst = df_inl["_dst"].apply(
            lambda u: urlparse(u).netloc.lower() in site_hosts if u.startswith("http") else True
        )
        df_int_links = df_inl[is_internal_dst].copy()
        df_ext_links = df_inl[~is_internal_dst].copy()

        # Internal Links Returning 404 Errors
        inl_404 = []
        mask_404 = df_int_links["_sc"].between(400, 499) | df_int_links["_dst"].map(
            lambda u: 400 <= url_to_status.get(u, 0) < 500
        )
        for _, r in df_int_links[mask_404].drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            sug = suggest_parent_redirect(r["_dst"], live_200_urls)
            fix_txt = f"Update link to: {sug}" if sug else "Remove broken link or update to live page"
            inl_404.append([r["_src"], r["_anc"], r["_dst"], fix_txt])
        record_check(
            "Internal Links to 404 pages",
            inl_404,
            ["Internal Links Returning 404 Errors"],
            "Found {count} internal link(s) pointing to 404 pages (see 'Internal Links to 404 pages' tab).",
        )

        # Internal Links That Are Redirecting
        inl_red = []
        mask_red = df_int_links["_sc"].between(300, 399) | df_int_links["_dst"].isin(redirect_map.keys())
        for _, r in df_int_links[mask_red].drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            true_dest = redirect_map[r["_dst"]][0] if r["_dst"] in redirect_map else ""
            inl_red.append([r["_src"], r["_anc"], r["_dst"], true_dest])
        record_check(
            "Internal Links to redirecting p",
            inl_red,
            ["Internal Links That Are Redirecting"],
            "Found {count} internal link(s) pointing to redirecting URLs (see 'Internal Links to redirecting p' tab).",
        )

        # External Links That Are Redirecting / Broken (Links to External pages tab)
        ext_red = []
        mask_ext_issue = (df_ext_links["_sc"] >= 300) | df_ext_links["_dst"].isin(redirect_map.keys())
        for _, r in df_ext_links[mask_ext_issue].drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            true_dest = redirect_map[r["_dst"]][0] if r["_dst"] in redirect_map else f"Status: {r['_sc']}"
            ext_red.append([r["_src"], r["_anc"], r["_dst"], true_dest])
        tab_data["Links to External pages"] = ext_red

        # Internal Links That Are Non-Canonical
        inl_noncanon = []
        for _, r in df_int_links.drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            dst = r["_dst"]
            st_c = r["_sc"] or url_to_status.get(dst, 0)
            canon = url_to_canonical.get(dst, "")
            if st_c == 200 and canon and canon != dst:
                inl_noncanon.append([r["_src"], r["_anc"], dst, canon])
        record_check(
            "InLinks to non-canonical pages",
            inl_noncanon,
            ["Internal Links That Are Non-Canonical"],
            "Found {count} internal link(s) pointing to non-canonical URLs (see 'InLinks to non-canonical pages' tab).",
        )

        # Internal Links That Are Unsecure (HTTP)
        inl_http = []
        for _, r in df_int_links[df_int_links["_dst"].str.startswith("http://")].drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            https_ver = "https://" + r["_dst"][len("http://"):]
            inl_http.append([r["_src"], r["_anc"], r["_dst"], https_ver])
        record_check(
            "InLinks to HTTP URLs",
            inl_http,
            ["Internal Links That Are Unsecure"],
            "Found {count} internal link(s) pointing to unsecure HTTP URLs (see 'InLinks to HTTP URLs' tab).",
        )

        # Nofollow / Noindex Internal Links
        inl_nofollow = []
        mask_nf = (df_int_links["_fol"] == "false") | df_int_links["_rel"].str.contains("nofollow", na=False)
        for _, r in df_int_links[mask_nf].drop_duplicates(subset=["_src", "_dst", "_anc"]).iterrows():
            inl_nofollow.append([r["_src"], r["_dst"], r["_anc"], r["_alt"]])
        record_check(
            "Nofollow internal links",
            inl_nofollow,
            ["Noindex internal links"],
            "Found {count} internal link(s) with nofollow attributes (see 'Nofollow internal links' tab).",
        )

        # Footer Links
        if pos_col:
            footer_dsts = sorted(df_int_links.loc[df_int_links["_pos"] == "footer", "_dst"].unique().tolist())
            footer_rows = [[u] for u in footer_dsts]
            record_check(
                "Footer links",
                footer_rows,
                ["Internal Link Issues | Footer links"],
                "Extracted {count} unique internal footer link destination(s) for review (see 'Footer links' tab).",
            )

    # --------------------------------------------------------------------------
    # 4. IMAGE AUDIT ENGINE (COMBINES IMAGES_ALL.CSV + ALL_INLINKS.CSV)
    # --------------------------------------------------------------------------
    img_status_map: dict[str, int] = {}
    img_size_map: dict[str, int] = {}

    if not df_images.empty:
        i_url_col = find_col(df_images, ["Address", "Image URL", "URL", "Destination"])
        i_sc_col = find_col(df_images, ["Status Code"])
        i_sz_col = find_col(df_images, ["Size (Bytes)", "Size"])
        for _, r in df_images.iterrows():
            u = str(r[i_url_col]).strip() if i_url_col and pd.notna(r[i_url_col]) else ""
            if u:
                if i_sc_col and pd.notna(r[i_sc_col]):
                    img_status_map[u] = int(pd.to_numeric(r[i_sc_col], errors="coerce") or 0)
                if i_sz_col and pd.notna(r[i_sz_col]):
                    img_size_map[u] = int(pd.to_numeric(r[i_sz_col], errors="coerce") or 0)

    # Fallback to internal_all lookup maps if images_all didn't have status/size
    for u, st_c in url_to_status.items():
        img_status_map.setdefault(u, st_c)
    for u, sz in url_to_size.items():
        img_size_map.setdefault(u, sz)

    # Check if user uploaded "All Image Inlinks" directly into df_images
    if df_img_inlinks.empty and not df_images.empty and find_col(df_images, ["Source"]):
        s_c = find_col(df_images, ["Source"])
        d_c = find_col(df_images, ["Destination", "Address", "Image URL"])
        a_c = find_col(df_images, ["Alt Text", "Alt"])
        sc_c = find_col(df_images, ["Status Code"])
        df_img_inlinks = pd.DataFrame({
            "_src": clean_str_series(df_images[s_c]),
            "_dst": clean_str_series(df_images[d_c]) if d_c else "",
            "_alt": clean_str_series(df_images[a_c]) if a_c else "",
            "_sc": pd.to_numeric(df_images[sc_c], errors="coerce").fillna(0).astype(int) if sc_c else 0,
        })

    if not df_img_inlinks.empty or not df_images.empty:
        # Broken 404 Images
        broken_imgs = []
        if not df_img_inlinks.empty:
            for _, r in df_img_inlinks.drop_duplicates(subset=["_src", "_dst"]).iterrows():
                sc = r["_sc"] or img_status_map.get(r["_dst"], 200)
                if sc >= 400:
                    broken_imgs.append([r["_src"], r["_dst"], "Replace broken image asset or remove broken <img> tag"])
        else:
            for u, sc in img_status_map.items():
                if sc >= 400 and re.search(r"\.(jpg|jpeg|png|webp|gif|svg)(\?|$)", u, re.I):
                    broken_imgs.append(["[See Inlinks in SF]", u, "Replace broken image asset or remove broken <img> tag"])
        record_check(
            "Broken 404 Images",
            broken_imgs,
            ["Broken 404 Images"],
            "Found {count} broken image reference(s) (see 'Broken 404 Images' tab).",
        )

        # Alt Attributes (Requires Image Inlinks from all_inlinks.csv or image_inlinks.csv)
        alt_rows = []
        if not df_img_inlinks.empty:
            for _, r in df_img_inlinks[df_img_inlinks["_alt"] == ""].drop_duplicates(subset=["_src", "_dst"]).iterrows():
                sc = r["_sc"] or img_status_map.get(r["_dst"], 200)
                if sc == 200:
                    alt_rows.append([r["_src"], r["_dst"], "", suggest_alt_from_filename(r["_dst"])])
            record_check(
                "Alt Attributes",
                alt_rows,
                ["Image Optimization | ALT Attributes"],
                "Found {count} image instance(s) missing Alt attributes (see 'Alt Attributes' tab).",
            )

        # Non-Descriptive File Name
        all_unique_imgs = set(img_status_map.keys()) | (set(df_img_inlinks["_dst"].unique()) if not df_img_inlinks.empty else set())
        non_desc_rows = []
        bad_name_re = re.compile(
            r"^(img|dsc|screenshot|screen[-_]?shot|image|photo|pic|untitled|asset|banner|slide|\d+|[_-])[0-9a-z_-]*\.(jpg|jpeg|png|webp|gif|svg)$",
            re.I,
        )
        for u in sorted(all_unique_imgs):
            fname = urlparse(u).path.split("/")[-1]
            if fname and bad_name_re.match(fname):
                non_desc_rows.append([u, "Rename with descriptive, hyphenated keyword filename"])
        tab_data["Non-Descriptive File Name"] = non_desc_rows

        # Image File Size (> 1MB / 1,048,576 bytes per template row 69 remark)
        large_img_rows = []
        if not df_img_inlinks.empty:
            seen_large = set()
            for _, r in df_img_inlinks.iterrows():
                if img_size_map.get(r["_dst"], 0) > 1_048_576 and r["_dst"] not in seen_large:
                    seen_large.add(r["_dst"])
                    large_img_rows.append([f"{r['_src']} ({r['_dst']})"])
        else:
            for u, sz in img_size_map.items():
                if sz > 1_048_576:
                    large_img_rows.append([u])
        record_check(
            "Image File Size",
            large_img_rows,
            ["Image File Size"],
            "Found {count} image(s) exceeding 1MB (1024 KB) in file size (see 'Image File Size' tab).",
        )

        # Products With 1 or No Image
        prod_urls = [u for u in live_200_urls if re.search(r"/(product|products)/[^/]+", u, re.I)]
        if prod_urls and not df_img_inlinks.empty:
            img_counts_per_page = df_img_inlinks.groupby("_src")["_dst"].nunique().to_dict()
            low_img_prods = [[u] for u in sorted(prod_urls) if img_counts_per_page.get(u, 0) <= 1]
            record_check(
                "Products With 1 or No Image",
                low_img_prods,
                ["Products With One or No Image"],
                "Found {count} product page(s) with 1 or 0 images (see 'Products With 1 or No Image' tab).",
            )

    # --------------------------------------------------------------------------
    # 5. OPTIONAL: GSC EXCLUDED / PAGES EXPORT
    # --------------------------------------------------------------------------
    if not df_gsc.empty:
        g_url_col = find_col(df_gsc, ["URL", "Page", "Address"])
        g_reason_col = find_col(df_gsc, ["Reason", "Status", "Issue", "Category"])
        if g_url_col and g_reason_col:
            df_g = df_gsc.copy()
            df_g["_u"] = clean_str_series(df_g[g_url_col])
            df_g["_r"] = clean_str_series(df_g[g_reason_col])

            soft_404_rows, gsc_excl_rows = [], []
            c_404 = c_dup_sub = c_dup_nosel = c_crawled = 0

            for _, r in df_g.iterrows():
                u, reason = r["_u"], r["_r"]
                r_low = reason.lower()
                if "soft 404" in r_low:
                    soft_404_rows.append([u, suggest_parent_redirect(u, live_200_urls)])
                else:
                    fix = "Review canonicalization, indexability, and internal linking."
                    if "not found (404)" in r_low or "404" in r_low:
                        c_404 += 1
                        fix = "Implement 301 redirect if valuable backlinks/traffic exist, or remove internal links."
                    elif "submitted url not selected as canonical" in r_low:
                        c_dup_sub += 1
                        fix = "Remove non-canonical URL from XML Sitemap and align canonical tags."
                    elif "duplicate without user-selected canonical" in r_low:
                        c_dup_nosel += 1
                        fix = "Add explicit self-referencing or consolidated rel='canonical' tag."
                    elif "crawled - currently not indexed" in r_low or "crawled currently not indexed" in r_low:
                        c_crawled += 1
                        fix = "Improve content depth/uniqueness and add contextual internal links."
                    gsc_excl_rows.append([u, reason, fix])

            record_check(
                "Soft 404 Errors",
                soft_404_rows,
                ["Internal Pages | Soft 404 Errors"],
                "Found {count} Soft 404 URL(s) in GSC export (see 'Soft 404 Errors' tab).",
            )
            tab_data["Google Search Console Excluded "] = gsc_excl_rows
            master_updates["GSC Excluded Pages | Not found (404)"] = (ISSUE_YES, f"GSC reports {c_404} 'Not found (404)' URL(s).") if c_404 else (ISSUE_NO, "0 GSC 404 exclusions.")
            master_updates["Duplicate, submitted URLs not selected as canonical"] = (ISSUE_YES, f"GSC reports {c_dup_sub} submitted URL(s) not selected as canonical.") if c_dup_sub else (ISSUE_NO, "0 submitted duplicate canonical issues.")
            master_updates["Duplicate without a user-selected canonical"] = (ISSUE_YES, f"GSC reports {c_dup_nosel} duplicate URL(s) without user-selected canonical.") if c_dup_nosel else (ISSUE_NO, "0 unselected canonical duplicates.")
            master_updates["Crawled currently not indexed"] = (ISSUE_YES, f"GSC reports {c_crawled} 'Crawled - currently not indexed' URL(s).") if c_crawled else (ISSUE_NO, "0 'Crawled - currently not indexed' URLs.")

    # --------------------------------------------------------------------------
    # 6. OPTIONAL: BACKLINKS EXPORT
    # --------------------------------------------------------------------------
    if not df_backlinks.empty:
        bl_target_col = find_col(df_backlinks, ["Target URL", "Target", "Destination URL", "URL", "Link To"])
        bl_dump = [df_backlinks.columns.tolist()] + df_backlinks.head(5000).fillna("").values.tolist()
        tab_data["Backlink profile analysis"] = bl_dump

        if bl_target_col and url_to_status:
            bl_404_urls = sorted({
                u for u in clean_str_series(df_backlinks[bl_target_col]).unique()
                if 400 <= url_to_status.get(u, 0) < 500
            })
            bl_404_rows = [[u, suggest_parent_redirect(u, live_200_urls)] for u in bl_404_urls]
            record_check(
                "Backlinks to non-existing pages",
                bl_404_rows,
                ["Backlinks Returning 404 Errors"],
                "Found {count} broken 404 target URL(s) receiving external backlinks (see 'Backlinks to non-existing pages' tab).",
            )

    return tab_data, master_updates


# ==============================================================================
# OPENPYXL WORKBOOK INJECTOR (DYNAMIC ITEM MATCHING & FORMATTING PRESERVATION)
# ==============================================================================
ROW_1_HEADER_TABS = {"Redirects to missing pages", "InLinks to HTTP URLs"}
OVERWRITE_ROW_2_TABS = {"Chains Redirects", "Backlink profile analysis"}


def compile_workbook(
    template_bytes: bytes,
    tab_data: dict[str, list[list]],
    master_updates: dict[str, tuple[str, str]],
) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))

    # 1. Dynamically Match ITEM Names on Master Sheet ('Copy of Initial Site Analysis |')
    master_name = "Copy of Initial Site Analysis |"
    if master_name in wb.sheetnames:
        ws_master = wb[master_name]
        for r_idx in range(2, ws_master.max_row + 1):
            item_val = ws_master.cell(row=r_idx, column=2).value
            if not item_val:
                continue
            item_str = str(item_val).strip().lower()
            for match_key, (status_flag, remark_text) in master_updates.items():
                if match_key.lower() in item_str:
                    ws_master.cell(row=r_idx, column=3, value=status_flag)
                    existing_remark = ws_master.cell(row=r_idx, column=7).value
                    if existing_remark:
                        ws_master.cell(
                            row=r_idx,
                            column=7,
                            value=f"[AUTO-AUDIT]: {remark_text}\n\n{existing_remark}",
                        )
                    else:
                        ws_master.cell(row=r_idx, column=7, value=f"[AUTO-AUDIT]: {remark_text}")
                    break

    # 2. Populate Evidence Tabs
    for sheet_name, rows in tab_data.items():
        if not rows or sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        if sheet_name in OVERWRITE_ROW_2_TABS or sheet_name in ROW_1_HEADER_TABS:
            start_row = 2
        else:
            start_row = 3

        for r_offset, row_vals in enumerate(rows):
            target_r = start_row + r_offset
            for c_idx, val in enumerate(row_vals, start=1):
                if isinstance(val, str):
                    val = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", val)
                ws.cell(row=target_r, column=c_idx, value=val)

    output_stream = io.BytesIO()
    wb.save(output_stream)
    return output_stream.getvalue()


# ==============================================================================
# STREAMLIT UI
# ==============================================================================
with st.sidebar:
    st.header("1. Template File (Required)")
    template_file = st.file_uploader(
        "Initial Site Analysis Template (.xlsx)",
        type=["xlsx"],
        help="Upload your agency's blank 44-tab Initial Site Analysis .xlsx template.",
    )

    st.header("2. Screaming Frog Exports")
    f_internal = st.file_uploader("internal_all.csv (Required)", type=["csv"])
    f_inlinks = st.file_uploader("all_inlinks.csv (Recommended)", type=["csv"])
    f_redirects = st.file_uploader("all_redirects.csv (Recommended)", type=["csv"])
    f_images = st.file_uploader("images_all.csv (Recommended)", type=["csv"])

    st.header("3. Optional Exports")
    f_gsc = st.file_uploader("GSC Pages / Excluded (.csv)", type=["csv"])
    f_backlinks = st.file_uploader("Backlinks Export (.csv)", type=["csv"])

if st.button("🚀 Compile Initial Site Analysis Workbook", type="primary", use_container_width=True):
    if not template_file or not f_internal:
        st.error("Please upload both your **Template (.xlsx)** and **internal_all.csv** to run the compiler.")
    else:
        with st.spinner("Parsing crawl data, resolving redirect chains, and compiling 44 tabs..."):
            df_int = load_csv_robust(f_internal)
            df_inl = load_csv_robust(f_inlinks)
            df_red = load_csv_robust(f_redirects)
            df_img = load_csv_robust(f_images)
            df_gsc = load_csv_robust(f_gsc)
            df_bl = load_csv_robust(f_backlinks)

            tab_data, master_updates = run_audit_engine(
                df_int, df_inl, df_red, df_img, df_gsc, df_bl
            )

            compiled_bytes = compile_workbook(
                template_file.getvalue(), tab_data, master_updates
            )

        st.success("✅ Initial Site Analysis v2.0 compiled! Review the summary below and download your workbook.")

        issues_flagged = sum(1 for v in master_updates.values() if v[0] == "✘Issues found")
        passed_checks = sum(1 for v in master_updates.values() if v[0] == "✔No issues")
        tabs_populated = sum(1 for rows in tab_data.values() if len(rows) > 0)

        c1, c2, c3 = st.columns(3)
        c1.metric("Checklist Items Flagged (✘)", issues_flagged)
        c2.metric("Checklist Items Passed (✔)", passed_checks)
        c3.metric("Evidence Tabs Populated", tabs_populated)

        summary_rows = [
            {"Evidence Tab": tab, "Rows Populated": len(rows)}
            for tab, rows in sorted(tab_data.items())
        ]
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        st.download_button(
            label="📥 Download Completed Initial Site Analysis (.xlsx)",
            data=compiled_bytes,
            file_name="Completed_Initial_Site_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
