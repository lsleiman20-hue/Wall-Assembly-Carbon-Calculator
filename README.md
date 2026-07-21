# Wall Assembly Carbon Calculator

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Upload this folder's contents to a GitHub repository.
2. In Streamlit Community Cloud, create a new app from that repository.
3. Set the main file path to `app.py`.
4. Deploy.

## Main workflow

Each new wall option begins with these fixed layer groups:

- Exterior Cladding
- Attachment System
- Insulation
- Stud and Framing
- Moisture and Air Control
- Sheathing
- Other

For every layer, select `None`, a preset material, or `Custom EPD`. Use **Add layer** inside a category only when the assembly needs another material in that group.


## kgCO₂e/kg mass inputs

For materials declared in kgCO₂e/kg, choose one of two simple inputs:

1. **Enter total mass (kg):** carbon = EPD GWP × total mass.
2. **Use grammage (kg/m²):** mass = wall area in m² × grammage, then carbon = EPD GWP × mass.

Spacing is not applied to this general mass method. Studs, rails, ties, and other repeated elements keep their dedicated calculation methods when needed.

## Interface overlap fix

This version removes the global CSS rule that forced Roboto onto Streamlit's Material Symbols icon font. It also includes dedicated styling for sidebar buttons, downloads, the JSON uploader, expanders, select boxes, and the material-library data grid.
