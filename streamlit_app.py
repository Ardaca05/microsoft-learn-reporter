import io
import re
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title='Microsoft Learn Reporter', page_icon='🎓', layout='wide')

PROFILE_API = 'https://learn.microsoft.com/api/profiles/{username}/newactivity?page=1&pageSize=1&locale=es-es'
ACHIEVEMENTS_API = 'https://learn.microsoft.com/api/achievements/user/{user_id}?locale=es-es'
HEADERS = {'User-Agent': 'Mozilla/5.0'}

st.title('🎓 Microsoft Learn Reporter')
st.caption('Consulta individual y masiva de Learning Paths públicos completados en Microsoft Learn.')


def clean_text(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    return str(value).strip()


def extract_username(value: str) -> str:
    value = clean_text(value)
    if not value:
        return ''
    match = re.search(r'/users/([^/?#]+)/?', value, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return value.strip('/ ')


def request_json(url, timeout=30, retries=2):
    last_error = None
    for attempt in range(retries + 1):
        try:
            response = requests.get(url, timeout=timeout, headers=HEADERS)
            if response.status_code == 429 and attempt < retries:
                time.sleep(2 * (attempt + 1))
                continue
            response.raise_for_status()
            return response.json()
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise last_error


def get_user_id(username: str) -> str:
    data = request_json(PROFILE_API.format(username=username), timeout=20)
    user_id = data.get('userId')
    if not user_id:
        raise ValueError('Microsoft Learn no devolvió un Learn User ID para este perfil.')
    return user_id


def get_learning_paths(user_id: str):
    data = request_json(ACHIEVEMENTS_API.format(user_id=user_id), timeout=30)
    achievements = data.get('achievements', [])
    return [a for a in achievements if str(a.get('category', '')).lower() == 'learningpaths']


def normalize_date(value):
    if not value:
        return pd.NaT
    try:
        return pd.to_datetime(value, utc=True).tz_convert(None)
    except Exception:
        return pd.NaT


def full_learn_url(relative_url):
    relative_url = clean_text(relative_url)
    if relative_url.startswith('/'):
        return 'https://learn.microsoft.com' + relative_url
    return relative_url


def path_rows(paths, employee_id='', email='', profile='', username='', user_id=''):
    rows = []
    for item in paths:
        rows.append({
            'ID Colaborador': employee_id,
            'Correo corporativo': email,
            'Perfil Microsoft Learn': profile or username,
            'Usuario Learn': username,
            'Learn User ID': user_id,
            'Learning Path': item.get('title', ''),
            'Fecha de finalización': normalize_date(item.get('grantedOn')),
            'URL Learning Path': full_learn_url(item.get('url', '')),
            'Type ID': item.get('typeId', ''),
        })
    return rows


def style_workbook(writer):
    header_fill = PatternFill('solid', fgColor='0B5CAB')
    header_font = Font(color='FFFFFF', bold=True)
    for ws in writer.book.worksheets:
        ws.freeze_panes = 'A2'
        ws.sheet_view.showGridLines = False
        if ws.max_row >= 1:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(vertical='center')
        for col in ws.columns:
            letter = col[0].column_letter
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 55)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, datetime):
                    cell.number_format = 'dd/mm/yyyy'


def make_excel(detail_df, summary_df, errors_df=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        detail_df.to_excel(writer, index=False, sheet_name='Learning Paths')
        summary_df.to_excel(writer, index=False, sheet_name='Resumen')
        if errors_df is not None:
            errors_df.to_excel(writer, index=False, sheet_name='Errores')
        style_workbook(writer)
    output.seek(0)
    return output.getvalue()


def make_template():
    template = pd.DataFrame([
        {
            'ID Colaborador': '000001',
            'Correo corporativo': 'usuario@empresa.com',
            'Perfil Microsoft Learn': 'https://learn.microsoft.com/es-es/users/usuario-1234/'
        },
        {
            'ID Colaborador': '',
            'Correo corporativo': '',
            'Perfil Microsoft Learn': 'otro-usuario-5678'
        },
    ])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template.to_excel(writer, index=False, sheet_name='Usuarios')
        style_workbook(writer)
    output.seek(0)
    return output.getvalue()


def display_table(df):
    shown = df.copy()
    if 'Fecha de finalización' in shown.columns:
        shown['Fecha de finalización'] = pd.to_datetime(shown['Fecha de finalización'], errors='coerce').dt.strftime('%d/%m/%Y')
    st.dataframe(shown, use_container_width=True, hide_index=True)


tab_individual, tab_mass = st.tabs(['🔎 Consulta individual', '📊 Consulta masiva'])

with tab_individual:
    st.subheader('Consulta individual')
    st.write('Ingresa un ID interno si deseas conservarlo en el reporte y pega la URL o username del perfil de Microsoft Learn.')
    c1, c2 = st.columns([1, 2])
    employee_id = c1.text_input('ID Colaborador (opcional)', key='individual_id', placeholder='Ej. 123456')
    profile = c2.text_input('Perfil Microsoft Learn *', key='individual_profile', placeholder='https://learn.microsoft.com/es-es/users/usuario-1234/')

    if st.button('Consultar Microsoft Learn', type='primary', key='individual_search'):
        username = extract_username(profile)
        if not username:
            st.warning('Ingresa una URL o username de Microsoft Learn.')
        else:
            try:
                with st.spinner('Consultando Microsoft Learn...'):
                    user_id = get_user_id(username)
                    paths = get_learning_paths(user_id)
                    rows = path_rows(paths, employee_id=clean_text(employee_id), profile=clean_text(profile), username=username, user_id=user_id)
                    detail = pd.DataFrame(rows)
                    if not detail.empty:
                        detail = detail.sort_values('Fecha de finalización', ascending=False, na_position='last').reset_index(drop=True)
                    last_date = detail['Fecha de finalización'].max() if not detail.empty else pd.NaT
                    summary = pd.DataFrame([{
                        'ID Colaborador': clean_text(employee_id),
                        'Perfil Microsoft Learn': clean_text(profile),
                        'Usuario Learn': username,
                        'Learn User ID': user_id,
                        'Learning Paths completados': len(detail),
                        'Última finalización': last_date,
                    }])

                st.success(f'Perfil encontrado. Se detectaron {len(detail)} Learning Paths.')
                m1, m2 = st.columns(2)
                m1.metric('Learning Paths', len(detail))
                m2.text_input('Learn User ID detectado', value=user_id, disabled=True, key='detected_id')

                if detail.empty:
                    st.info('No se encontraron Learning Paths públicos para este perfil.')
                else:
                    display_table(detail[['ID Colaborador', 'Learning Path', 'Fecha de finalización', 'URL Learning Path']])

                excel = make_excel(detail, summary)
                st.download_button(
                    'Descargar Excel', excel,
                    file_name=f'Microsoft_Learn_{username}_{datetime.now().strftime("%Y%m%d")}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key='individual_download'
                )
            except Exception as exc:
                st.error(f'No se pudo procesar el perfil: {exc}')

with tab_mass:
    st.subheader('Consulta masiva')
    st.write('Carga un Excel con una fila por colaborador. Solo **Perfil Microsoft Learn** es obligatorio; ID y correo son opcionales.')
    st.download_button(
        'Descargar plantilla Excel', make_template(),
        file_name='Plantilla_Microsoft_Learn_Reporter.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key='template_download'
    )
    uploaded = st.file_uploader('Cargar archivo Excel', type=['xlsx'], key='mass_upload')

    if uploaded is not None:
        try:
            source = pd.read_excel(uploaded, dtype=str).fillna('')
            required = 'Perfil Microsoft Learn'
            optional_cols = ['ID Colaborador', 'Correo corporativo']
            if required not in source.columns:
                st.error(f'El archivo debe contener la columna obligatoria: {required}')
            else:
                for col in optional_cols:
                    if col not in source.columns:
                        source[col] = ''
                source = source[['ID Colaborador', 'Correo corporativo', 'Perfil Microsoft Learn']]
                source = source[source['Perfil Microsoft Learn'].astype(str).str.strip() != ''].reset_index(drop=True)
                st.info(f'{len(source)} perfiles listos para procesar.')
                st.dataframe(source.head(10), use_container_width=True, hide_index=True)

                if st.button('Procesar perfiles', type='primary', key='mass_process'):
                    all_rows = []
                    summaries = []
                    errors = []
                    progress = st.progress(0, text='Preparando...')
                    total = len(source)

                    for idx, row in source.iterrows():
                        employee_id = clean_text(row['ID Colaborador'])
                        email = clean_text(row['Correo corporativo'])
                        profile_value = clean_text(row['Perfil Microsoft Learn'])
                        username = extract_username(profile_value)
                        progress.progress((idx + 1) / max(total, 1), text=f'Procesando {idx + 1} de {total}: {username or profile_value}')
                        try:
                            user_id = get_user_id(username)
                            paths = get_learning_paths(user_id)
                            rows = path_rows(paths, employee_id, email, profile_value, username, user_id)
                            all_rows.extend(rows)
                            dates = [normalize_date(p.get('grantedOn')) for p in paths]
                            dates = [d for d in dates if not pd.isna(d)]
                            summaries.append({
                                'ID Colaborador': employee_id,
                                'Correo corporativo': email,
                                'Perfil Microsoft Learn': profile_value,
                                'Usuario Learn': username,
                                'Learn User ID': user_id,
                                'Learning Paths completados': len(paths),
                                'Última finalización': max(dates) if dates else pd.NaT,
                                'Estado': 'OK',
                            })
                        except Exception as exc:
                            errors.append({
                                'ID Colaborador': employee_id,
                                'Correo corporativo': email,
                                'Perfil Microsoft Learn': profile_value,
                                'Usuario Learn': username,
                                'Error': str(exc),
                            })
                            summaries.append({
                                'ID Colaborador': employee_id,
                                'Correo corporativo': email,
                                'Perfil Microsoft Learn': profile_value,
                                'Usuario Learn': username,
                                'Learn User ID': '',
                                'Learning Paths completados': 0,
                                'Última finalización': pd.NaT,
                                'Estado': 'Error',
                            })

                    progress.empty()
                    detail = pd.DataFrame(all_rows, columns=[
                        'ID Colaborador', 'Correo corporativo', 'Perfil Microsoft Learn', 'Usuario Learn', 'Learn User ID',
                        'Learning Path', 'Fecha de finalización', 'URL Learning Path', 'Type ID'
                    ])
                    if not detail.empty:
                        detail = detail.sort_values(['ID Colaborador', 'Correo corporativo', 'Fecha de finalización'], ascending=[True, True, False], na_position='last').reset_index(drop=True)
                    summary = pd.DataFrame(summaries)
                    errors_df = pd.DataFrame(errors, columns=['ID Colaborador', 'Correo corporativo', 'Perfil Microsoft Learn', 'Usuario Learn', 'Error'])

                    st.session_state['mass_result'] = (detail, summary, errors_df)
        except Exception as exc:
            st.error(f'No se pudo leer el archivo: {exc}')

    if 'mass_result' in st.session_state:
        detail, summary, errors_df = st.session_state['mass_result']
        st.success(f'Proceso terminado: {len(summary)} perfiles procesados, {len(detail)} Learning Paths encontrados y {len(errors_df)} errores.')
        m1, m2, m3 = st.columns(3)
        m1.metric('Perfiles procesados', len(summary))
        m2.metric('Learning Paths', len(detail))
        m3.metric('Errores', len(errors_df))

        st.markdown('#### Resumen')
        display_table(summary)
        st.markdown('#### Learning Paths')
        display_table(detail)
        if not errors_df.empty:
            st.markdown('#### Errores')
            st.dataframe(errors_df, use_container_width=True, hide_index=True)

        excel = make_excel(detail, summary, errors_df)
        st.download_button(
            'Descargar consolidado Excel', excel,
            file_name=f'Microsoft_Learn_Consolidado_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='mass_download'
        )

st.divider()
st.caption('Consulta información pública expuesta por los endpoints utilizados por los perfiles de Microsoft Learn. Antes de un uso corporativo masivo, valida las políticas y condiciones aplicables de Microsoft.')
