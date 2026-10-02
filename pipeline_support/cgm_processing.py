from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from statistics import mean
from typing import Callable, cast

from openpyxl import load_workbook
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq


PARTICIPANT_IDS = ['01', '02', '03', '04', '05']
INTERVAL_MINUTES = 5
INTERVAL = timedelta(minutes=INTERVAL_MINUTES)

# Keep the pilot switch in the module because it belongs to CGM extraction, not to
# the downstream donation joins. Normal processing always uses the full workbook.
USE_PILOT_WINDOW = False
PILOT_START_UTC = datetime(2024, 3, 5, 0, 0, tzinfo=timezone.utc)
PILOT_DAYS = 3
PILOT_END_UTC = PILOT_START_UTC + timedelta(days=PILOT_DAYS)


def default_fail_validation(message: str) -> None:
    """Fallback used when the module is run outside DataProcessing.ipynb."""
    raise RuntimeError(message)


def utc_from_milliseconds(value: str | int | float | None) -> datetime | None:
    """Convert the workbook timestamp column to an aware UTC datetime."""
    if value in (None, ''):
        return None
    return datetime.fromtimestamp(float(value) / 1000.0, tz=timezone.utc)


def timestamp_key(value: str | int | float | None) -> int | None:
    """Build a rounded timestamp key for detecting duplicate BG-check rows."""
    if value in (None, ''):
        return None
    return round(float(value))


def public_id(participant: str) -> str:
    return str(participant).split('_', 1)[0]


def donation_schema(donation_columns: list[str], schema: dict[str, list[str]]) -> pa.Schema:
    """Create the exact Arrow schema expected by the donation table."""
    string_columns = set(schema['string_columns'])
    boolean_columns = set(schema['boolean_columns'])
    fields = []
    for column in donation_columns:
        if column in string_columns:
            fields.append(pa.field(column, pa.string()))
        elif column in boolean_columns:
            fields.append(pa.field(column, pa.bool_()))
        else:
            fields.append(pa.field(column, pa.float64()))
    return pa.schema(fields)


def empty_donation_row(
    participant: str,
    timestamp: datetime,
    donation_columns: list[str],
    schema: dict[str, list[str]],
) -> dict[str, object]:
    """Create one donation-schema row with only CGM backbone fields populated."""
    string_columns = set(schema['string_columns'])
    boolean_columns = set(schema['boolean_columns'])
    row: dict[str, object] = {}
    for column in donation_columns:
        if column in string_columns or column in boolean_columns:
            row[column] = None
        else:
            row[column] = float('nan')
    row['date'] = timestamp.strftime('%Y-%m-%d %H:%M:%S')
    row['id'] = public_id(participant)
    return row


def rows_to_table(
    rows: list[dict[str, object]],
    donation_columns: list[str],
    schema: dict[str, list[str]],
) -> pa.Table:
    """Convert row dictionaries to a schema-fixed Arrow table."""
    columns = {column: [row.get(column) for row in rows] for column in donation_columns}
    return pa.table(columns, schema=donation_schema(donation_columns, schema))


def iter_excel_records(workbook_path: Path):
    """Yield raw workbook rows with participant and parsed timestamp metadata."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        for participant_id in PARTICIPANT_IDS:
            sheet = workbook[participant_id]
            rows = sheet.iter_rows(values_only=True)
            # The workbook sheets share a header row; empty or non-text headers are
            # ignored because they cannot identify a usable event field.
            headers = tuple(value.strip() if isinstance(value, str) else '' for value in next(rows))
            for values in rows:
                record = dict(zip(headers, values, strict=False))
                record['participant'] = participant_id
                unix_timestamp = record.get('UNIX GTM+0')
                record['timestamp_utc'] = utc_from_milliseconds(unix_timestamp)
                record['timestamp_key'] = timestamp_key(unix_timestamp)
                yield record
    finally:
        workbook.close()


def bg_check_timestamps_by_participant(workbook_path: Path) -> dict[str, set[int]]:
    """Collect BG-check timestamps that duplicate workbook CGM entries."""
    timestamps: dict[str, set[int]] = {participant_id: set() for participant_id in PARTICIPANT_IDS}
    for record in iter_excel_records(workbook_path):
        participant = str(record['participant'])
        origin = str(record.get('Origen') or '').strip().lower()
        event_type = str(record.get('eventType') or '').strip()
        timestamp = record.get('timestamp_key')
        if origin == 'treatments' and event_type == 'BG Check' and isinstance(timestamp, int):
            timestamps.setdefault(participant, set()).add(timestamp)
    return timestamps


def in_selected_window(timestamp: datetime) -> bool:
    """Apply the optional pilot window without changing the main extraction code."""
    if not USE_PILOT_WINDOW:
        return True
    timestamp_utc = timestamp.astimezone(timezone.utc)
    return PILOT_START_UTC <= timestamp_utc < PILOT_END_UTC


def load_cgm_readings(workbook_path: Path) -> dict[str, list[tuple[datetime, float]]]:
    """Read only true CGM entry rows and remove duplicated BG-check timestamps."""
    readings: dict[str, list[tuple[datetime, float]]] = {participant_id: [] for participant_id in PARTICIPANT_IDS}
    bg_check_timestamps = bg_check_timestamps_by_participant(workbook_path)
    for record in iter_excel_records(workbook_path):
        timestamp = record.get('timestamp_utc')
        timestamp_for_dedup = record.get('timestamp_key')
        origin = str(record.get('Origen') or '').strip().lower()
        participant = str(record['participant'])
        duplicates_bg_check = timestamp_for_dedup in bg_check_timestamps.get(participant, set())
        # Nightscout-style exports can contain BG Check treatment rows with the
        # same timestamp as an entry row. Those are treatment duplicates, not CGM
        # observations, so they are filtered before building the strict grid.
        if (
            origin == 'entries'
            and isinstance(timestamp, datetime)
            and in_selected_window(timestamp)
            and record.get('sgv') is not None
            and not duplicates_bg_check
        ):
            readings.setdefault(participant, []).append((timestamp.astimezone(timezone.utc), float(record['sgv'])))
    for participant in readings:
        readings[participant].sort(key=lambda item: item[0])
    return readings


def grid_index(anchor: datetime, timestamp: datetime) -> int:
    """Return the nearest five-minute grid index relative to a subject anchor."""
    elapsed = (timestamp - anchor).total_seconds()
    return max(0, round(elapsed / INTERVAL.total_seconds()))


def split_consecutive_runs(readings: list[tuple[datetime, float]]) -> list[list[tuple[datetime, float]]]:
    """Split readings when the gap no longer behaves like a five-minute CGM cadence."""
    if not readings:
        return []

    max_gap_error = INTERVAL.total_seconds() / 2
    runs: list[list[tuple[datetime, float]]] = [[readings[0]]]
    for previous, current in zip(readings, readings[1:]):
        previous_timestamp = previous[0]
        current_timestamp = current[0]
        gap_seconds = (current_timestamp - previous_timestamp).total_seconds()
        # A run can drift a little around five minutes, but a larger gap means the
        # next reading should establish its own local placement instead of shifting
        # every later reading forward.
        if abs(gap_seconds - INTERVAL.total_seconds()) <= max_gap_error:
            runs[-1].append(current)
        else:
            runs.append([current])
    return runs


def extend_rows(
    rows: list[dict[str, object]],
    participant: str,
    anchor: datetime,
    target_index: int,
    donation_columns: list[str],
    schema: dict[str, list[str]],
) -> None:
    """Extend a participant grid until the requested index exists."""
    while len(rows) <= target_index:
        rows.append(empty_donation_row(participant, anchor + len(rows) * INTERVAL, donation_columns, schema))


def assign_to_grid_row(
    rows: list[dict[str, object]],
    assignments: dict[int, dict[str, object]],
    resolved_conflicts: list[dict[str, object]],
    participant: str,
    anchor: datetime,
    target_index: int,
    timestamp: datetime,
    value: float,
    donation_columns: list[str],
    schema: dict[str, list[str]],
    fail_validation: Callable[[str], None],
) -> None:
    """Assign one observed CGM value to one strict-grid row.

    Two readings can occasionally compete for the same grid row after the workbook
    duplicate filtering. The closest reading wins and the dropped reading is kept
    in the conflict audit instead of being silently discarded.
    """
    extend_rows(rows, participant, anchor, target_index, donation_columns, schema)
    target_time = anchor + target_index * INTERVAL
    delta_seconds = abs((target_time - timestamp).total_seconds())
    existing = assignments.get(target_index)
    if existing is None:
        rows[target_index]['CGM'] = value
        rows[target_index]['_source_timestamp_utc'] = timestamp
        assignments[target_index] = {
            'timestamp': timestamp,
            'value': value,
            'delta_seconds': delta_seconds,
        }
        return

    existing_delta_seconds = cast(float, existing['delta_seconds'])
    existing_timestamp = cast(datetime, existing['timestamp'])
    if not isinstance(existing_timestamp, datetime):
        fail_validation('Existing CGM assignment has an invalid timestamp.')

    keep_new = delta_seconds < existing_delta_seconds
    kept = {'timestamp': timestamp, 'value': value, 'delta_seconds': delta_seconds} if keep_new else existing
    dropped = existing if keep_new else {'timestamp': timestamp, 'value': value, 'delta_seconds': delta_seconds}
    dropped_timestamp = cast(datetime, dropped['timestamp'])
    if not isinstance(dropped_timestamp, datetime):
        fail_validation('Dropped CGM assignment has an invalid timestamp.')
    resolved_conflicts.append({
        'id': public_id(participant),
        'grid_index': target_index,
        'grid_time': target_time.isoformat(),
        'kept_timestamp': kept['timestamp'].isoformat(),
        'kept_CGM': kept['value'],
        'kept_delta_seconds': round(kept['delta_seconds'], 3),
        'dropped_timestamp': dropped_timestamp.isoformat(),
        'dropped_CGM': dropped['value'],
        'dropped_delta_seconds': round(cast(float, dropped['delta_seconds']), 3),
        'seconds_between_readings': round(abs((timestamp - existing_timestamp).total_seconds()), 3),
    })
    if keep_new:
        rows[target_index]['CGM'] = value
        rows[target_index]['_source_timestamp_utc'] = timestamp
        assignments[target_index] = {
            'timestamp': timestamp,
            'value': value,
            'delta_seconds': delta_seconds,
        }


def build_participant_rows(
    participant: str,
    readings: list[tuple[datetime, float]],
    donation_columns: list[str],
    schema: dict[str, list[str]],
    fail_validation: Callable[[str], None],
    resolved_conflicts: list[dict[str, object]] | None = None,
) -> list[dict[str, object]]:
    """Build a strict five-minute CGM grid for one participant."""
    if not readings:
        return []
    if resolved_conflicts is None:
        resolved_conflicts = []

    anchor = readings[0][0]
    last_timestamp = readings[-1][0]
    last_index = grid_index(anchor, last_timestamp)
    rows = [
        empty_donation_row(participant, anchor + index * INTERVAL, donation_columns, schema)
        for index in range(last_index + 1)
    ]

    max_anchor_delta = INTERVAL.total_seconds() / 2
    assignments: dict[int, dict[str, object]] = {}
    skipped_readings: list[tuple[datetime, float, str]] = []

    for run in split_consecutive_runs(readings):
        run_start_timestamp = run[0][0]
        run_start_index = grid_index(anchor, run_start_timestamp)
        run_start_time = anchor + run_start_index * INTERVAL
        anchor_delta_seconds = abs((run_start_time - run_start_timestamp).total_seconds())
        # Only the first reading of each consecutive run is snapped to the nearest
        # grid row. The remaining readings in that run advance by one row each,
        # preserving the observed cadence without repeatedly re-snapping drift.
        if anchor_delta_seconds > max_anchor_delta:
            skipped_readings.extend((timestamp, value, 'run start farther than half interval') for timestamp, value in run)
            continue

        for offset, (timestamp, value) in enumerate(run):
            target_index = run_start_index + offset
            assign_to_grid_row(
                rows,
                assignments,
                resolved_conflicts,
                participant,
                anchor,
                target_index,
                timestamp,
                value,
                donation_columns,
                schema,
                fail_validation,
            )

    if skipped_readings:
        examples = ', '.join(
            f'{timestamp.isoformat()} ({reason})'
            for timestamp, _, reason in skipped_readings[:3]
        )
        fail_validation(
            f'{participant}: {len(skipped_readings)} CGM readings could not be assigned to the strict grid. '
            f'Examples: {examples}'
        )

    return rows


def build_rows_by_participant(
    workbook_path: Path,
    donation_columns: list[str],
    schema: dict[str, list[str]],
    fail_validation: Callable[[str], None],
) -> tuple[dict[str, list[dict[str, object]]], list[dict[str, object]]]:
    """Build all participant grids and return the conflict audit."""
    readings_by_participant = load_cgm_readings(workbook_path)
    resolved_conflicts: list[dict[str, object]] = []
    rows_by_participant = {
        participant: build_participant_rows(participant, readings, donation_columns, schema, fail_validation, resolved_conflicts)
        for participant, readings in readings_by_participant.items()
        if readings
    }
    return rows_by_participant, resolved_conflicts


def create_cgm_parquet(
    project_root: Path,
    raw_dir: Path,
    in_process_dir: Path,
    schema: dict[str, list[str]],
    donation_columns: list[str],
    fail_validation: Callable[[str], None] = default_fail_validation,
    display: Callable[[object], None] | None = None,
) -> pd.DataFrame:
    """Create the CGM backbone Parquet and its validation summary.

    DataProcessing.ipynb injects paths, schema, display and fail_validation so this
    module uses the same runtime configuration and notebook-friendly validation
    messages as the main pipeline. The historical CGM notebook remains as a
    reference, but this function is now the executable implementation.
    """
    raw_workbook = raw_dir / 'All patients CGM and Events.xlsx'
    cgm_output_path = in_process_dir / 'cgm.parquet'
    summary_output_path = in_process_dir / 'cgm_validation_summary.json'

    print(f'Project root: {project_root}')
    print(f'Raw workbook: {raw_workbook}')
    print(f'CGM output: {cgm_output_path}')
    if USE_PILOT_WINDOW:
        print(f'Pilot window UTC: {PILOT_START_UTC.isoformat()} to {PILOT_END_UTC.isoformat()}')

    raw_cgm_readings = load_cgm_readings(raw_workbook)
    rows_by_participant, resolved_grid_conflicts = build_rows_by_participant(
        raw_workbook,
        donation_columns,
        schema,
        fail_validation,
    )
    if display is not None and resolved_grid_conflicts:
        display(pd.DataFrame(resolved_grid_conflicts))

    all_rows = [
        row
        for participant_rows in rows_by_participant.values()
        for row in participant_rows
    ]

    boundary_fix_counts = {
        'cgm_39_to_40': 0,
        'cgm_401_to_400': 0,
    }
    for row in all_rows:
        # The accepted donation range is inclusive. Values one unit outside the
        # boundary are known workbook artifacts and are corrected explicitly; more
        # distant out-of-range values remain validation failures below.
        if row['CGM'] == 39:
            row['CGM'] = 40.0
            boundary_fix_counts['cgm_39_to_40'] += 1
        elif row['CGM'] == 401:
            row['CGM'] = 400.0
            boundary_fix_counts['cgm_401_to_400'] += 1

    cgm_table = rows_to_table(all_rows, donation_columns, schema)
    pq.write_table(cgm_table, cgm_output_path)

    summary_rows = []
    for participant, rows in rows_by_participant.items():
        participant_id = public_id(participant)
        cgm_values = [row['CGM'] for row in rows]
        non_null = sum(1 for value in cgm_values if value == value)
        nan_count = len(rows) - non_null
        observed_rows = [row for row in rows if row['CGM'] == row['CGM']]
        deltas = []
        for row in observed_rows:
            real_timestamp = cast(datetime, row['_source_timestamp_utc'])
            date_value = row['date']
            if not isinstance(date_value, str):
                fail_validation('CGM grid row has an invalid date value.')
            date_text = cast(str, date_value)
            assigned = datetime.strptime(date_text, '%Y-%m-%d %H:%M:%S').replace(tzinfo=timezone.utc)
            deltas.append((assigned - real_timestamp).total_seconds())
        abs_deltas = sorted(abs(delta) for delta in deltas)
        p99_abs = abs_deltas[round((len(abs_deltas) - 1) * 0.99)]
        resolved_for_participant = sum(
            1 for conflict in resolved_grid_conflicts
            if conflict['id'] == participant_id
        )

        summary_rows.append({
            'id': participant_id,
            'rows': len(rows),
            'raw CGM readings': len(raw_cgm_readings[participant]),
            'assigned CGM': non_null,
            'resolved conflicts': resolved_for_participant,
            'missing CGM': nan_count,
            'mean delta seconds': round(mean(deltas), 3),
            'min delta seconds': round(min(deltas), 3),
            'max delta seconds': round(max(deltas), 3),
            'p99 abs delta seconds': round(p99_abs, 3),
            'max abs delta seconds': round(max(abs_deltas), 3),
            'abs delta >= 150s': sum(1 for delta in deltas if abs(delta) >= 150),
        })

    print(f'Consolidated CGM Parquet: {cgm_output_path.relative_to(project_root)}')
    print(f'Total rows: {len(all_rows)}')
    print(f'Participants: {len(summary_rows)}')
    print(f'Boundary fixes: {boundary_fix_counts}')
    if display is not None:
        display(pd.DataFrame(summary_rows))

    written_table = pq.read_table(cgm_output_path)
    ids = written_table.column('id').to_pylist()
    dates = written_table.column('date').to_pylist()
    cgm_values = written_table.column('CGM').to_pylist()

    strict_grid_ok = True
    bad_grid_examples = []
    for participant_id in sorted(set(ids)):
        timestamps = [
            datetime.strptime(date, '%Y-%m-%d %H:%M:%S').replace(tzinfo=timezone.utc)
            for row_id, date in zip(ids, dates)
            if row_id == participant_id
        ]
        deltas = [
            int((right - left).total_seconds())
            for left, right in zip(timestamps, timestamps[1:])
        ]
        bad_deltas = sorted(set(delta for delta in deltas if delta != INTERVAL.total_seconds()))
        if bad_deltas:
            strict_grid_ok = False
            bad_grid_examples.append(f'{participant_id}: {bad_deltas[:3]}')

    out_of_range_cgm = [
        value for value in cgm_values
        if value == value and not 40 <= value <= 400
    ]

    validation_summary = {
        'status': 'CGM backbone OK' if strict_grid_ok and not out_of_range_cgm else 'CGM backbone needs review',
        'parquet_files': 1,
        'total_rows': written_table.num_rows,
        'participants': len(set(ids)),
        'assigned_cgm': sum(row['assigned CGM'] for row in summary_rows),
        'resolved_conflicts': len(resolved_grid_conflicts),
        'conflicts_resolved': True,
        'missing_cgm': sum(row['missing CGM'] for row in summary_rows),
        'boundary_fixes_39_to_40': boundary_fix_counts['cgm_39_to_40'],
        'boundary_fixes_401_to_400': boundary_fix_counts['cgm_401_to_400'],
        'out_of_range_cgm_after_fixes': len(out_of_range_cgm),
        'columns': len(donation_columns),
        'column_order_is_correct': written_table.column_names == donation_columns,
        'strict_5_minute_grid': strict_grid_ok,
        'bad_grid_examples': '; '.join(bad_grid_examples),
        'output': str(cgm_output_path.relative_to(project_root)),
    }

    with summary_output_path.open('w', encoding='utf-8') as summary_file:
        json.dump(validation_summary, summary_file, indent=2)

    if display is not None:
        display(pd.DataFrame([validation_summary]))

    return pd.read_parquet(cgm_output_path)
