# Wall Assembly Embodied Carbon Calculator

## Files

- `app.py` — complete Streamlit application
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — Streamlit theme configuration

## Deploy on Streamlit Community Cloud

1. Upload all files and the `.streamlit` folder to your GitHub repository.
2. Commit the changes.
3. Deploy `app.py` through Streamlit Community Cloud.
4. Existing deployed apps should update automatically after the GitHub commit.

## Changes in this version

- Removed project JSON upload and download controls.
- Kept the formatted Excel export.
- Changed emitted and net carbon intensity to `kgCO₂e/m²` throughout the app and Excel export.
- Added persistent browser storage for reusable custom materials.
- Added Add, Edit, and Remove controls for custom materials. Preset materials remain read-only.
- Added a Definitions page explaining A1–A3, the selected scope, stored/biogenic carbon, net carbon, carbon intensity, and net intensity.
- Updated the wall-layer order to:
  1. Exterior Cladding
  2. Attachment System
  3. Moisture and Air Control
  4. Sheathing
  5. Insulation
  6. Stud and Framing
  7. Interior Finishes
- Improved the open-sidebar collapse-arrow contrast.

## Custom-material storage

Custom materials are stored in the current browser using local storage. They remain after a page refresh or after returning to the app in the same browser. They are not automatically shared with a different browser or device, and clearing browser site data will remove them.


## Automatic formula selection

Custom materials no longer require a calculation-method selection. The app derives the formula from the EPD declared unit:

- kgCO₂e/m² → wall area
- kgCO₂e/m³ → wall area × thickness
- kgCO₂e/kg → total mass or wall area × grammage
- kgCO₂e/m → total installed linear length
- kgCO₂e/unit → installed item count
- kgCO₂e/m² at R/RSI → wall area adjusted to the required R-value

Preset wood and steel framing products retain their specialized spacing and gauge/size formulas.


## Definitions update

The Definitions page now includes:
- RSI and R-value meanings and conversion formulas
- The EPD declared unit and where to locate it in an EPD
- Revised stored/biogenic carbon wording
