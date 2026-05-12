# CNCC Exa Scale

CNCC Exa Scale is a Python project for standardizing activity workbook data and
calculating quarterly person-time statistics.

## Project Layout

- `standardize.py`: standardizes the source activity workbook.
- `quarterly_stats.py`: runs standardization and then calculates quarterly
  person-time statistics.
- `src/cncc_exa_scale/flows/`: workflow orchestration.
- `src/cncc_exa_scale/modules/`: reusable processing modules.
- `config.yaml`: runtime configuration with environment variable overrides.

## Inputs And Outputs

Runtime input files are read from `input/` by default. Generated files are
written to `output/`, and logs are written to `log/`.

These runtime directories are intentionally ignored by Git.

## Usage

Install the runtime dependencies in your Python environment:

```powershell
pip install pandas openpyxl pyyaml
```

Place the source workbook in `input/`, then run:

```powershell
python standardize.py
```

To generate quarterly person-time statistics:

```powershell
python quarterly_stats.py
```

## Configuration

The default configuration is in `config.yaml`.

Key settings:

- `app.input_path`: input directory, defaults to `input`.
- `app.output_dir`: output directory, defaults to `output`.
- `flows.standardize_activity_info.input_file`: source workbook name.
- `flows.quarterly_person_times.years`: comma-separated comparison years.

Local environment overrides can be placed in `common.env`; that file should not
be committed.
