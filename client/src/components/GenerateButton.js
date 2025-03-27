import React from 'react';
import { Button } from '@mui/material';

export default function GenerateButton({ onClick }) {
  return (
    <Button variant="contained" color="success" onClick={onClick}>
      엑셀 생성
    </Button>
  );
}
