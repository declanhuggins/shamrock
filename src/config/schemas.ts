// Tab schemas and canonical header definitions used during setup.

namespace Schemas {
  // Minimal machine headers for key tables. Display headers default to machine headers if not provided.
  export const FRONTEND_TABS: Types.TabSchema[] = [
    { name: 'FAQs', machineHeaders: ['faq'] },
    { name: 'Dashboard', machineHeaders: ['metric', 'value'] },
    {
      name: 'Leadership',
      machineHeaders: ['last_name', 'first_name', 'rank', 'role', 'reports_to', 'email', 'office_phone', 'cell_phone', 'office_location'],
    },
    {
      name: 'Directory',
      machineHeaders: [
        'last_name',
        'first_name',
        'as_year',
        'flight',
        'squadron',
        'university',
        'email',
        'phone_display',
        'dorm',
        'home_town',
        'home_state',
        'class_year',
        'dob',
        'cip_broad_area',
        'cip_code',
        'desired_assigned_afsc',
        'flight_path_status',
        'photo_link',
        'notes',
      ],
    },
    {
      name: 'Attendance',
      machineHeaders: ['last_name', 'first_name', 'as_year', 'flight', 'squadron', 'overall_attendance_pct', 'llab_attendance_pct'],
    },
    {
      name: 'Data Legend',
      machineHeaders: [
        'as_year_options',
        'flight_options',
        'squadron_options',
        'university_options',
        'dorm_options',
        'home_state_options',
        'cip_broad_area_options',
        'afsc_options',
        'flight_path_status_options',
        'attendance_code_options',
      ],
      displayHeaders: [
        'AS Year Options',
        'Flight Options',
        'Squadron Options',
        'University Options',
        'Dorm Options',
        'Home State Options',
        'CIP Broad Area Options',
        'AFSC Options',
        'Flight Path Status Options',
        'Attendance Code Options',
      ],
    },
  ];

  export const BACKEND_TABS: Types.TabSchema[] = [
    {
      name: 'Directory Backend',
      machineHeaders: ['source', 'last_name', 'first_name', 'as_year', 'class_year', 'flight', 'squadron', 'university', 'email', 'phone', 'dorm', 'home_town', 'home_state', 'dob', 'cip_broad_area', 'cip_code', 'desired_assigned_afsc', 'flight_path_status', 'photo_link', 'notes'],
    },
    {
      name: 'Leadership Backend',
      machineHeaders: ['last_name', 'first_name', 'rank', 'role', 'reports_to', 'email', 'office_phone', 'cell_phone', 'office_location'],
    },
    {
      name: 'Events Backend',
      machineHeaders: ['event_id', 'term', 'training_week', 'event_type', 'display_name', 'attendance_column_label', 'expected_group', 'flight_scope', 'status', 'start_datetime', 'end_datetime', 'location', 'notes', 'created_at', 'created_by'],
    },
    {
      name: 'Excusals Backend',
      machineHeaders: ['request_id', 'event', 'email', 'last_name', 'first_name', 'flight', 'squadron', 'status', 'decision', 'decided_by', 'decided_at', 'attendance_effect', 'submitted_at', 'last_updated_at', 'notes'],
    },
    {
      name: 'Attendance Backend',
      machineHeaders: ['submission_id', 'submitted_at', 'event', 'email', 'name', 'flight', 'cadets'],
    },
    {
      name: 'Audit Backend',
      machineHeaders: ['audit_id', 'timestamp', 'actor_email', 'role', 'action', 'target_sheet', 'target_table', 'target_key', 'target_range', 'event_id', 'request_id', 'old_value', 'new_value', 'result', 'reason', 'notes', 'source', 'version', 'run_id'],
    },
    {
      name: 'Data Legend',
      machineHeaders: [
        'as_year_options',
        'flight_options',
        'squadron_options',
        'university_options',
        'dorm_options',
        'home_state_options',
        'cip_broad_area_options',
        'afsc_options',
        'flight_path_status_options',
        'attendance_code_options',
      ],
      displayHeaders: [
        'AS Year Options',
        'Flight Options',
        'Squadron Options',
        'University Options',
        'Dorm Options',
        'Home State Options',
        'CIP Broad Area Options',
        'AFSC Options',
        'Flight Path Status Options',
        'Attendance Code Options',
      ],
    },
  ];
}
