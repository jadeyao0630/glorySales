import React, { useRef, useState } from 'react';
import axios from 'axios';
interface Data {
    // Define the structure of your data here
  }
  const ip='192.168.10.213'
  const port="5555"
const FileUploadComponent: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
    const [message, setMessage] = useState<string>('');
    const [year, setYear] = useState<number>(new Date().getFullYear());
    const [month, setMonth] = useState<number>(new Date().getMonth()+1);
    const inputYearRef = useRef<HTMLInputElement>(null);
    const inputMonthRef = useRef<HTMLInputElement>(null);
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files.length > 0) {
      setFile(event.target.files[0]);
    }
  };
  const handleYearInputChange = () => {
    if (inputYearRef.current) {
      const inputValue = parseInt(inputYearRef.current.value);
      if (!isNaN(inputValue)) {
        setYear(inputValue);
      }
    }
  };
  const handleMonthInputChange = () => {
    if (inputMonthRef.current) {
      const inputValue = parseInt(inputMonthRef.current.value);
      if (!isNaN(inputValue)) {
        setMonth(inputValue);
      }
    }
  };
  const handleDownload = async () => {
    try {
        // Make a GET request with parameters
        const _y=inputYearRef.current?.value || '2024';
        const _m=inputMonthRef.current?.value || '0';
        const response = await axios.get(`http://${ip}:${port}/projects`, {
            responseType: 'blob',
          params: {
            year: _y,
            month: _m,
          }
        });

        const downloadUrl = window.URL.createObjectURL(response.data);

      // Create a link element and click it to trigger the download
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.setAttribute('download', `template_${_y}${_m==='0'?'':'_'+_m}.xlsx`); // Set desired filename here
      document.body.appendChild(link);
      link.click();

      // Cleanup: remove the link and revoke the URL object
      link.parentNode?.removeChild(link);
      window.URL.revokeObjectURL(downloadUrl);
      } catch (error) {
        console.error('Error fetching data:', error);
      }
    };
  
  const handleUpload = async () => {
    if (!file) {
      alert('请选择一个文件。');
      return;
    }

    const formData = new FormData();
    formData.append('excel', file);

    try {
      const response = await axios.post(`http://${ip}:${port}/upload`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      });
      
      // Handle response from backend
      setMessage(response.data);
      console.log('Response from backend:', response.data);
    } catch (error) {
      console.error('Error uploading file:', error);
    }
  };

  return (
    <div>
      <div className='main_title'>
        营销目标计划平台
      </div>
        <div>
        <input type="number" ref={inputYearRef} placeholder='年' onChange={handleYearInputChange} value={year}/>
        <label className='input_label'>年</label>
        <input type="number" ref={inputMonthRef} placeholder='月' onChange={handleMonthInputChange} value={month}/>
        <label className='input_label'>月</label>
        <button onClick={handleDownload}>下载模板</button>
        <label className='input_label'>设置月为零则下载全年模板</label>
        </div>
        <div>
        <input type="file" accept=".xlsx" onChange={handleFileChange} />
        <button onClick={handleUpload}>上传数据</button>
        </div>
        
        <div dangerouslySetInnerHTML={{ __html: message }} >
        </div>
    </div>
  );
};

export default FileUploadComponent;