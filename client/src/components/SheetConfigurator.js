import React from 'react';
import { FormControlLabel, Checkbox, Typography } from '@mui/material';

export default function SheetConfigurator({ headers, selected, setSelected }) {
  const handleChange = (name) => {
    const next = selected.includes(name)
      ? selected.filter((h) => h !== name)
      : [...selected, name];
    setSelected(next);
  };

  const uniqueGroups = [...new Set(headers.filter(Boolean))];

  return (
    <>
      <Typography variant="h6">시트로 분리할 그룹 선택</Typography>
      {uniqueGroups.map((group) => (
        <FormControlLabel
          key={group}
          control={
            <Checkbox
              checked={selected.includes(group)}
              onChange={() => handleChange(group)}
            />
          }
          label={group}
        />
      ))}
    </>
  );
}
