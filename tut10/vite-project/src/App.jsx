import { useState } from 'react'
import reactLogo from './assets/react.svg'
import viteLogo from '/vite.svg'
import './App.css'
import FileUploadDownload from './FileUploader'

function App() {
  const [file, setFile] = useState(0)

  return (
    <>
      <FileUploadDownload/>
    </>
  )
}

export default App
