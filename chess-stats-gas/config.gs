var PROJECT = {
  sheetTitle: 'Games'
};

// Each header can map from a data field via `from`, or compute via `formula`.
// Use placeholders in formulas: ${row}, ${col(key)} to reference cells.
var HEADERS = [
  { key: 'source',        title: 'Source',          from: 'source',        formula: null,                                   visible: true,  width: 72,  color: '#E3F2FD' },
  { key: 'username',      title: 'User',            from: 'username',      formula: null,                                   visible: true,  width: 100, color: '#E3F2FD' },
  { key: 'url',           title: 'URL',             from: 'url',           formula: null,                                   visible: true,  width: 260, color: '#E3F2FD' },
  { key: 'end_time',      title: 'End Time (UTC)',  from: 'end_time',      formula: null,                                   visible: true,  width: 150, color: '#E3F2FD' },
  { key: 'opponent',      title: 'Opponent',        from: 'opponent',      formula: null,                                   visible: true,  width: 120, color: '#E3F2FD' },
  { key: 'user_color',    title: 'Color',           from: 'user_color',    formula: null,                                   visible: true,  width: 65,  color: '#E3F2FD' },
  { key: 'result_simple', title: 'Result (W/D/L)',  from: 'result_simple', formula: null,                                   visible: true,  width: 110, color: '#E3F2FD' },
  { key: 'time_control',  title: 'Time Control',    from: 'time_control',  formula: null,                                   visible: true,  width: 110, color: '#E3F2FD' },
  { key: 'eco',           title: 'ECO',             from: 'eco',           formula: null,                                   visible: false, width: 60,  color: '#E3F2FD' },
  { key: 'termination',   title: 'Termination',     from: 'termination',   formula: null,                                   visible: false, width: 160, color: '#E3F2FD' },
  { key: 'pgn',           title: 'PGN',             from: 'pgn',           formula: null,                                   visible: false, width: 500, color: '#E3F2FD' }
];