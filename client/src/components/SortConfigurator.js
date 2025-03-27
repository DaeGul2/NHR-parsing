import React from 'react';
import {
  Typography,
  Select,
  MenuItem,
  FormControl,
  InputLabel
} from '@mui/material';

export default function SortConfigurator({ group, sortKeys, setSortKeys }) {
  const handleChange = (e) => {
    setSortKeys({ ...sortKeys, [group]: e.target.value });
  };

  return (
    <FormControl fullWidth>
      <InputLabel>정렬 기준</InputLabel>
      <Select
        value={sortKeys[group] || ''}
        label="정렬 기준"
        onChange={handleChange}
      >
        <MenuItem value="">정렬 안 함</MenuItem>
        <MenuItem value="오름차순">오름차순</MenuItem>
        <MenuItem value="내림차순">내림차순</MenuItem>
      </Select>
      <Typography variant="caption">({group}) 기준</Typography>
    </FormControl>
  );
}
