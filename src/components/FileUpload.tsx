import React, { useCallback } from 'react';
import './FileUpload.css';

interface FileUploadProps {
  onFileUpload: (file: File) => void;
  accept?: string;
}

const FileUpload: React.FC<FileUploadProps> = ({ 
  onFileUpload, 
  accept = '.xlsx,.xls,.csv' 
}) => {
  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      const file = e.dataTransfer.files[0];
      if (file) {
        onFileUpload(file);
      }
    },
    [onFileUpload]
  );

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) {
        onFileUpload(file);
      }
    },
    [onFileUpload]
  );

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  return (
    <div
      className="file-upload"
      onDrop={handleDrop}
      onDragOver={handleDragOver}
    >
      <div className="upload-icon">📁</div>
      <p>Drag and drop an Excel file here</p>
      <p className="upload-or">or</p>
      <label className="upload-button">
        <input
          type="file"
          accept={accept}
          onChange={handleChange}
          style={{ display: 'none' }}
        />
        Browse Files
      </label>
      <p className="upload-hint">Supports: .xlsx, .xls, .csv</p>
    </div>
  );
};

export default FileUpload;
