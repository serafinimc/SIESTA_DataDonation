# DataDonation Processing
This project builds one consolidated DataDonation parquet from the raw project exports.

## CRITICAL
Raw data files must be requested from cecilia.serafini@ing.unlp.edu.ar and added to the rawdata folder.

## First-time setup

Create and activate a virtual environment from the repository root, then install the required packages:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run the pipeline

Open `DataProcessing.ipynb` and run it from top to bottom. The notebook validates each intermediate step before publishing the final parquet. To rerun only the final validation/export from existing intermediates, set `RUN_ONLY_FINAL_VALIDATION = True` in the first cell before running the notebook.

You can open the notebook from PyCharm or from any Jupyter-compatible interface using the virtual environment created above.

Expected final output:

```text
finalData/SIESTA.parquet
```

The raw files under `rawdata/` are not modified during processing. Intermediate files are written under `inProcess/`, including `inProcess/final_validation_audit.csv` with the same column-level summary shown in the final validation notebook table. The final schema-validated parquet is published under `finalData/`.
