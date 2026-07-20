# Wall Carbon Streamlit App

This project converts the connected Excel wall-assembly GWP calculator into a Streamlit web application.

## Files

- `app.py` — all formulas, material presets, interface, comparison charts, JSON saving, and Excel export.
- `requirements.txt` — Python packages Streamlit Cloud installs.
- `.streamlit/config.toml` — colors and interface theme.

## Run in a browser with Streamlit Community Cloud

1. Create a new GitHub repository.
2. Upload all files and folders from this project. Keep `.streamlit/config.toml` inside the `.streamlit` folder.
3. Open Streamlit Community Cloud and choose **Create app**.
4. Select the repository, branch, and `app.py`.
5. Deploy.

## Run locally (optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Important calculation note

For `Mass by Grammage`, choose **Adjust grammage for spacing? = Yes** only when the entered kg/m² is based on a reference spacing. The app multiplies mass by `reference spacing / actual spacing`.
