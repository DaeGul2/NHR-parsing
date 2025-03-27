import React from 'react';
import {
  FormGroup,
  FormControlLabel,
  Checkbox,
  Typography
} from '@mui/material';

export default function ColumnSelector({ group, headerRow, columns, setColumns }) {
  const groupIndices = headerRow
    .map((g, idx) => (g === group ? idx : null))
    .filter((v) => v !== null);

  const handleToggle = (idx) => {
    const exists = columns[group] || [];
    const next = exists.includes(idx)
      ? exists.filter((i) => i !== idx)
      : [...exists, idx];
    setColumns({ ...columns, [group]: next });
  };

  return (
    <>
      <Typography variant="h6">{group} 컬럼 선택</Typography>
      <FormGroup>
        {groupIndices.map((idx) => (
          <FormControlLabel
            key={idx}
            control={
              <Checkbox
                checked={columns[group]?.includes(idx) || false}
                onChange={() => handleToggle(idx)}
              />
            }
            label={`열 ${idx + 1}`}
          />
        ))}
      </FormGroup>
    </>
  );
}
