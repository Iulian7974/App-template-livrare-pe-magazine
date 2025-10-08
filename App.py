import io
import re
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generator template N-ERP", layout="wide")
st.title("Generator template N-ERP – foi per depozit & fișiere per depozit")

st.markdown("""
Încarcă **fișierul de exemplu** (.xlsx) cu coloanele: `Warehouse`, `Material Code`, `Quantity`, `New Price`.

Format foaie/template (fix):
- `Material Code`, `Quantity`, `Val. Type` = **A**, `New Price`
- `I/C New Price` (gol), `I/C New Cur.` (gol)
- `Plant` = **L402**
- `S / L` (gol), `Biz.ModelGroup` (gol), `Biz.Category` (gol)
""")

# Structura finală a fiecărei foi
TEMPLATE_COLS = [
    "Material Code",
    "Quantity",
    "Val. Type",
    "New Price",
    "I/C New Price",
    "I/C New Cur.",
    "Plant",
    "S / L",
    "Biz.ModelGroup",
    "Biz.Category",
]

def normalize_columns(df):
    """Normalizează denumirile (lower, fără spații duble) ca să acceptăm mici variații."""
    mapping = {}
    for c in df.columns:
        key = " ".join(str(c).strip().split()).lower()
        mapping[c] = key
    return df.rename(columns=mapping)

def find_column(df_norm, candidates):
    """Caută o coloană în df_norm după o listă de variante posibile (lower)."""
    for cand in candidates:
        if cand in df_norm.columns:
            return cand
    return None

def validate_and_extract(src_df):
    """Validează și extrage coloanele cerute, întorcând un DataFrame curat cu numele standard."""
    df_norm = normalize_columns(src_df)

    col_wh = find_column(df_norm, ["warehouse", "depozit"])
    col_mat = find_column(df_norm, ["material code", "material", "cod material", "material_code"])
    col_qty = find_column(df_norm, ["quantity", "qty", "cantitate"])
    col_price = find_column(df_norm, ["new price", "pret nou", "new_price"])

    missing = []
    if not col_wh: missing.append("Warehouse")
    if not col_mat: missing.append("Material Code")
    if not col_qty: missing.append("Quantity")
    if not col_price: missing.append("New Price")
    if missing:
        raise ValueError("Lipsesc coloanele obligatorii: " + ", ".join(missing))

    out = pd.DataFrame({
        "Warehouse": src_df[df_norm.columns[df_norm.columns.get_loc(col_wh)]],
        "Material Code": src_df[df_norm.columns[df_norm.columns.get_loc(col_mat)]].astype(str),
        "Quantity": pd.to_numeric(src_df[df_norm.columns[df_norm.columns.get_loc(col_qty)]], errors="coerce"),
        "New Price": pd.to_numeric(src_df[df_norm.columns[df_norm.columns.get_loc(col_price)]], errors="coerce"),
    })

    # curățare rânduri fără depozit/material
    out = out.dropna(subset=["Warehouse", "Material Code"])
    return out

def build_template_for_subset(df_subset):
    """Construiește DataFrame-ul în forma finală de template, pe baza subsetului (un depozit)."""
    out = pd.DataFrame(columns=TEMPLATE_COLS)
    out["Material Code"] = df_subset["Material Code"].astype(str)
    out["Quantity"] = df_subset["Quantity"]
    out["Val. Type"] = "A"
    out["New Price"] = df_subset["New Price"]
    out["I/C New Price"] = ""
    out["I/C New Cur."] = ""
    out["Plant"] = "L402"
    out["S / L"] = ""
    out["Biz.ModelGroup"] = ""
    out["Biz.Category"] = ""
    return out[TEMPLATE_COLS]

def sanitize_sheet_name(name):
    """
    Curăță denumirea foii ca să fie acceptată de Excel:
    - max 31 caractere
    - fără: : \ / ? * [ ]
    - înlocuiește spații multiple
    """
    s = str(name).strip()
    s = re.sub(r'[:\\/\?\*\[\]]', '-', s)
    s = re.sub(r'\s+', ' ', s)
    if len(s) > 31:
        s = s[:31]
    return s if s else "Sheet"

uploaded = st.file_uploader("Încarcă fișierul de exemplu (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        src_df = pd.read_excel(uploaded)
        clean_df = validate_and_extract(src_df)

        warehouses = pd.unique(clean_df["Warehouse"].dropna())
        # sortare: numerice înainte, apoi lexicografic
        def sort_key(x):
            sx = str(x)
            return (0, int(sx)) if sx.isdigit() else (1, sx)
        warehouses = sorted(warehouses, key=sort_key)

        st.success(f"Găsite {len(warehouses)} depozite.")
        st.caption(f"Exemple: {', '.join(map(str, warehouses[:10]))}{'…' if len(warehouses) > 10 else ''}")

        st.divider()
        st.subheader("Descărcări")

        col1, col2, col3 = st.columns(3)

        # 1) Un singur Excel cu foi separate
        with col1:
            st.markdown("**Workbook cu foi per depozit**")
            if st.button("Generează un singur Excel (toate depozitele)"):
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    used_names = set()
                    for wh in warehouses:
                        df_wh = clean_df[clean_df["Warehouse"] == wh].copy()
                        out_df = build_template_for_subset(df_wh)
                        sheet_name = sanitize_sheet_name(wh)
                        base = sheet_name
                        i = 1
                        while sheet_name in used_names:
                            suffix = f"_{i}"
                            sheet_name = (base[:(31-len(suffix))] + suffix) if len(base)+len(suffix) > 31 else base + suffix
                            i += 1
                        used_names.add(sheet_name)
                        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
                buffer.seek(0)
                st.download_button(
                    label="Descarcă warehouse_templates.xlsx",
                    data=buffer,
                    file_name="warehouse_templates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        # 2) ZIP cu fișiere per depozit
        with col2:
            st.markdown("**Fișiere separate (ZIP)**")
            if st.button("Generează ZIP cu fișiere per depozit"):
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for wh in warehouses:
                        df_wh = clean_df[clean_df["Warehouse"] == wh].copy()
                        out_df = build_template_for_subset(df_wh)
                        xbuf = io.BytesIO()
                        out_df.to_excel(xbuf, index=False)
                        xbuf.seek(0)
                        zf.writestr(f"Warehouse {wh}.xlsx", xbuf.read())
                zip_buf.seek(0)
                st.download_button(
                    label="Descarcă warehouse_templates.zip",
                    data=zip_buf,
                    file_name="warehouse_templates.zip",
                    mime="application/zip",
                )

        # 3) Un singur depozit (opțional)
        with col3:
            st.markdown("**Depozit individual (opțional)**")
            sel_wh = st.selectbox("Alege depozit", warehouses, key="single_wh")
            if st.button("Generează Excel pentru depozitul selectat"):
                df_wh = clean_df[clean_df["Warehouse"] == sel_wh].copy()
                out_df = build_template_for_subset(df_wh)
                buf = io.BytesIO()
                out_df.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button(
                    label=f"Descarcă Warehouse {sel_wh}.xlsx",
                    data=buf,
                    file_name=f"Warehouse {sel_wh}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        st.divider()
        st.subheader("Previzualizare (opțional)")
        if len(warehouses) > 0:
            sel_prev = st.selectbox("Alege depozitul pentru previzualizare", warehouses, key="preview_wh")
            df_prev = build_template_for_subset(clean_df[clean_df["Warehouse"] == sel_prev].copy())
            st.dataframe(df_prev.head(20))
        st.caption("Structura este fixă: Val. Type=A, Plant=L402; restul coloanelor rămân goale dar prezente.")

    except Exception as e:
        st.error(f"Eroare: {e}")
else:
    st.info("Încarcă fișierul de exemplu pentru a începe.")
