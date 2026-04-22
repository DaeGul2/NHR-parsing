import React, { useState, useCallback } from 'react';
import { Button, Box, Typography } from '@mui/material';

export default function FileUploader({ onUpload }) {
  const [dragging, setDragging] = useState(false);

  const handleChange = (e) => {
    const file = e.target.files?.[0];
    if (file) onUpload(file);
  };

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragging(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragging(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && /\.(xlsx|xls|csv)$/i.test(file.name)) onUpload(file);
  }, [onUpload]);

  return (
    <Box
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      sx={{
        border: '2px dashed',
        borderColor: dragging ? 'primary.main' : 'divider',
        borderRadius: 2,
        p: 2,
        textAlign: 'center',
        backgroundColor: dragging ? 'action.hover' : 'transparent',
        transition: 'all 0.2s',
        cursor: 'pointer',
      }}
    >
      <Button variant="contained" component="label">
        NHR 형식 엑셀 업로드
        <input type="file" hidden accept=".xlsx,.xls,.csv" onChange={handleChange} />
      </Button>
      <Typography variant="caption" display="block" sx={{ mt: 0.5, color: 'text.secondary' }}>
        또는 파일을 여기에 드래그
      </Typography>
    </Box>
  );
}
