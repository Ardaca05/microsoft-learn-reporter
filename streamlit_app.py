import io
import re
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Microsoft Learn Reporter", page_icon="🎓", layout="wide")

st.title("🎓 Microsoft Learn Reporter")
st.caption("Consulta los Learning Paths públicos completados desde un perfil de Microsoft Learn.")

PROFILE_API = "https://learn.microsoft.com/api/profiles/{username}/newactivity?page=1&pageSize=1&locale=es-es"
ACHIEVEMENTS_API = "https://learn.microsoft.com/api/achievements/user/{user_id}?locale=es-es"


def extract_username(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    match = re.search(r"/users/([^/?#]+)/?", value, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return value.strip("/ ")


def get_user_id(username: str) -> str:
    url = PROFILE_API.format(username=username)
    response = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    data = response.json()
    user_id = data.get("userId")
    if not user_id:
        raise ValueError("Microsoft Learn no devolvió un User ID para este perfil.")
    return user_id


def get_learning_paths(user_id: str):
    url = ACHIEVEMENTS_API.format(user_id=user_id)
    response = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    data = response.json()
    achievements = data.get("achievements", [])
    paths = [a for a in achievements if str(a.get("category", "")).lower() == "learningpaths"]
    return paths


def normalize_date(value):
    if not value:
        return None
    try:
        return pd.to_datetime(value, utc=True).tz_convert(None)
    except Exception:
        return value


def build_dataframe(paths):
    rows = []
    for item in paths:
        relative_url = item.get("url", "") or ""
        full_url = f"https://learn.microsoft.com{relative_url}" if relative_url.startswith("/") else relative_url
        rows.append({
            "Learning Path": item.get("title", ""),
            "Fecha de finalización": normalize_date(item.get("grantedOn")),
            "URL": full_url,
            "Type ID": item.get("typeId", ""),
            "User ID": item.get("userId", ""),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("Fecha de finalización", ascending=False, na_position="last").reset_index(drop=True)
    return df


def make_excel(df, username, user_id):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Learning Paths")
        resumen = pd.DataFrame([
            {"Perfil": username, "Learn User ID": user_id, "Learning Paths": len(df),
             "Última finalización": df["Fecha de finalización"].max() if not df.empty else None}
        ])
        resumen.to_excel(writer, index=False, sheet_name="Resumen")

        for sheet_name in ["Learning Paths", "Resumen"]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 60)
    output.seek(0)
    return output.getvalue()

profile = st.text_input(
    "URL o username de Microsoft Learn",
    value="alvaroadamincasapa-0791",
    placeholder="https://learn.microsoft.com/es-es/users/usuario-1234/",
)

if st.button("Consultar Microsoft Learn", type="primary", use_container_width=False):
    username = extract_username(profile)
    if not username:
        st.warning("Ingresa una URL o username de Microsoft Learn.")
    else:
        try:
            with st.spinner("Consultando Microsoft Learn..."):
                user_id = get_user_id(username)
                paths = get_learning_paths(user_id)
                df = build_dataframe(paths)

            st.success(f"Perfil encontrado. Se detectaron {len(df)} Learning Paths.")
            c1, c2 = st.columns(2)
            c1.metric("Learning Paths", len(df))
            c2.text_input("Learn User ID detectado", value=user_id, disabled=True)

            if df.empty:
                st.info("No se encontraron Learning Paths públicos para este perfil.")
            else:
                display_df = df.copy()
                if "Fecha de finalización" in display_df:
                    display_df["Fecha de finalización"] = pd.to_datetime(display_df["Fecha de finalización"], errors="coerce").dt.strftime("%d/%m/%Y")
                st.dataframe(display_df[["Learning Path", "Fecha de finalización", "URL"]], use_container_width=True, hide_index=True)

                excel_bytes = make_excel(df, username, user_id)
                st.download_button(
                    "Descargar Excel",
                    data=excel_bytes,
                    file_name=f"Microsoft_Learn_{username}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with st.expander("Detalle técnico"):
                st.write("Username:", username)
                st.write("Learn User ID:", user_id)
                st.write("Endpoint consultado:", ACHIEVEMENTS_API.format(user_id=user_id))

        except requests.HTTPError as exc:
            st.error(f"Microsoft Learn respondió con un error HTTP: {exc}")
        except requests.RequestException as exc:
            st.error(f"No se pudo conectar con Microsoft Learn: {exc}")
        except Exception as exc:
            st.error(f"No se pudo procesar el perfil: {exc}")

st.divider()
st.caption("POC. Consulta únicamente información expuesta por los endpoints públicos utilizados por el perfil de Microsoft Learn.")
