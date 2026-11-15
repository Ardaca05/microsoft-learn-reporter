# Microsoft Learn Reporter

Aplicación Streamlit con dos modos:

- **Consulta individual:** ID Colaborador opcional + Perfil Microsoft Learn obligatorio.
- **Consulta masiva:** carga XLSX con `ID Colaborador`, `Correo corporativo` y `Perfil Microsoft Learn`. Solo el perfil es obligatorio.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Community Cloud

Sube `streamlit_app.py`, `requirements.txt` y `README.md` al repositorio y configura `streamlit_app.py` como Main file path.
