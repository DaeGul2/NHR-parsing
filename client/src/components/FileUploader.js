import React from 'react';
import { Button } from '@mui/material';

export default function FileUploader({ onUpload }) {
  const handleChange = (e) => {
    const file = e.target.files?.[0];
    if (file) {
      onUpload(file);
    }
  };

  return (
    <Button variant="contained" component="label">
      엑셀 업로드
      <input type="file" hidden accept=".xlsx" onChange={handleChange} />
    </Button>
  );
}
