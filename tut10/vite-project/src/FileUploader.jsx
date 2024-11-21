import React, { useState } from 'react';

function FileUploadDownload() {
  const [file, setFile] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleFileUpload = async (event) => {
    const uploadedFile = event.target.files[0];

      setFile(uploadedFile);
      setError("");
      setLoading(true);

      const formData = new FormData();
      formData.append('file', uploadedFile);

      try {
        const response = await fetch('http://localhost:3000/upload', {
          method: 'POST',
          body: formData,
        });
        console.log("here")

        if (!response.ok) {
          throw new Error('File processing failed.');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.download = 'OutputFile.xlsx'; // Name for the downloaded file
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);

      } catch (err) {
        console.error('Error:', err);
        setError('An error occurred while processing the file.');
      } finally {
        setLoading(false);
      }

  };

  return (
    <div>
      <h1>IIT Patna Grader</h1>
      <h2>Upload the input File</h2>
      <input
        type="file"
        onChange={handleFileUpload}
        accept=".xlsx" 
      />
      <br />
      <br />
      {loading && <p>Processing your file, please wait...</p>}
      {error.length > 0 && <div>{error}</div>}
    </div>
  );
}

export default FileUploadDownload;
