import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import "../css/Form.css";
import { fetchProvinceList, filterKabupatenList } from "../Services/services";
import { motion } from "framer-motion";

function Form() {
  const [status, setStatus] = useState("");
  const [selectedProvince, setSelectedProvince] = useState("");
  const [selectedKabupaten, setSelectedKabupaten] = useState("");
  const [provinces, setProvinces] = useState([]);
  const [kabupatenList, setKabupatenList] = useState([]);

  const [FormData, setFormData] = useState({
    provincial: { lgName: "", email: "", phone: "" },
    kabupaten: { lgName: "", email: "", phone: "" },
  });

  const [files, setFiles] = useState({
    BridgeInventory: null,
    CODE_AN_Parameters: null,
    CODE_AN_UnitCostsPER: null,
    CODE_AN_UnitCostsPERUnpaved: null,
    CODE_AN_UnitCostsREH: null,
    CODE_AN_UnitCostsRIGID: null,
    CODE_AN_UnitCostsRM: null,
    CODE_AN_UnitCostsUPGUnpaved: null,
    CODE_AN_UnitCostsWidening: null,
    CODE_AN_WidthStandards: null,
    CulvertCondition: null,
    CulvertInventory: null,
    Link: null,
    RetainingWallInventory: null,
    RoadInventory: null,
    TrafficVolume: null,
    RoadCondition: null,
  });

  const [excelJson, setExcelJson] = useState({});
  const [output, setOutput] = useState(null);

  useEffect(() => {
    setProvinces(fetchProvinceList());
  }, []);

  const handleProvinceChange = useCallback(
    (e) => {
      const selectedLG = e.target.value;
      const province = provinces.find((p) => p.LG === selectedLG);
      if (!province) return;

      setSelectedProvince(selectedLG);
      setSelectedKabupaten("");
      setKabupatenList(filterKabupatenList(province.adm_code));
    },
    [provinces]
  );

  const handleInputChange = useCallback((e, type, field) => {
    const value = e.target.value;
    setFormData((prev) => ({
      ...prev,
      [type]: { ...prev[type], [field]: value },
    }));
  }, []);

  const handleFileChange = useCallback(async (e, key) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "buffer" });
      
      // Process all sheets in the workbook
      const allData = {};
      workbook.SheetNames.forEach(sheetName => {
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        allData[sheetName] = jsonData;
      });

      setFiles((prev) => ({ ...prev, [key]: file }));
      setExcelJson((prev) => ({ ...prev, [key]: allData[workbook.SheetNames[0]] }));
    } catch (error) {
      alert("Error reading Excel file. Please check the file format.");
      console.error(error);
    }
  }, []);

  const removeFile = useCallback((key) => {
    setFiles((prev) => ({ ...prev, [key]: null }));
    setExcelJson((prev) => {
      const updated = { ...prev };
      delete updated[key];
      return updated;
    });
  }, []);

  const validateInputs = () => {
    if (
      !status ||
      !selectedProvince ||
      (status === "kabupaten" && !selectedKabupaten)
    ) {
      alert("Please fill out all dropdown fields.");
      return false;
    }

    const { lgName, email, phone } = FormData[status];

    // LG Name validation
    if (!lgName.trim()) {
      alert("LG Name is required.");
      return false;
    }

    // Email validation
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    if (!emailRegex.test(email)) {
      alert("Please enter a valid email address (e.g., example@domain.com)");
      return false;
    }

    // Phone number validation (only digits, min 9 max 14 characters)
    const phoneRegex = /^[0-9]{9,14}$/;
    if (!phoneRegex.test(phone.replace(/^\+62/, ''))) {
      alert("Phone number must be 9 to 14 digits and contain only numbers.");
      return false;
    }

    return true;
  };

  const handleSubmit = async () => {
    if (!validateInputs()) return;

    const prependPrefix = (phone) => {
      if (!phone) return "";
      return phone.startsWith("+62") ? phone : `+62${phone.replace(/^0+/, "")}`;
    };

    const currentFormData = FormData[status];
    const preparedFormData = {
      status: status,
      selected_province: selectedProvince,
      selected_kabupaten: status === "kabupaten" ? selectedKabupaten : null,
      lg_name: currentFormData.lgName,
      email: currentFormData.email,
      phone: prependPrefix(currentFormData.phone)
    };
    
    // Create the data structure with FormData as an array
    const jsonData = {
      FormData: [preparedFormData]  // Wrap FormData in an array
    };

    // Add Excel data for each file
    Object.entries(excelJson).forEach(([key, data]) => {
      if (data && Array.isArray(data)) {
        jsonData[key] = data;
      }
    });

    console.log('Sending data:', jsonData); // Log the data being sent

    try {
      const response = await fetch('http://127.0.0.1:8000/api/upload-data/', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(jsonData),
      });

      const responseData = await response.json();
      console.log('Response from server:', responseData); // Log the complete response

      if (!response.ok) {
        // If there are validation errors, show them
        if (responseData.details && responseData.details.FormData) {
          const formDataErrors = responseData.details.FormData;
          if (formDataErrors.status === 'validation_error') {
            const errorMessages = Object.values(formDataErrors.errors)
              .map(error => error.errors)
              .join('\n');
            throw new Error(`Validation errors:\n${errorMessages}`);
          }
        }
        throw new Error(responseData.message || `HTTP error! status: ${response.status}`);
      }

      // Check if FormData was successfully processed
      if (responseData.FormData && responseData.FormData.status === 'validated') {
        alert('Data submitted successfully!');
      } else {
        console.warn('FormData processing status:', responseData.FormData);
        alert('Data was received but there might be an issue with processing. Check console for details.');
      }
    } catch (error) {
      console.error('Error details:', error);
      alert(`Error submitting data: ${error.message}`);
    }

    setOutput(jsonData);
  };

  const renderFormInputs = (type) => (
    <motion.div
      className="form-inputs"
      initial={{ opacity: 0, x: -30 }}
      animate={{ opacity: 1, x: 0 }}
      transition={{ duration: 0.4 }}
    >
      <h3 className="section-title">Details</h3>
      <div className="form-row">
  {/* LG Name */}
  <div className="input-group">
    <label className="floating-label">LG Name</label>
    <input
      type="text"
      placeholder="Enter LG Name"
      value={FormData[type].lgName}
      onChange={(e) => handleInputChange(e, type, "lgName")}
      className="input-field enhanced"
    />
  </div>

  {/* Email */}
  <div className="input-group">
    <label className="floating-label">Email</label>
    <input
      type="email"
      placeholder="Enter Email"
      value={FormData[type].email}
      onChange={(e) => handleInputChange(e, type, "email")}
      className="input-field enhanced"
    />
  </div>

  {/* Phone Number */}
  <div className="input-group">
    <label className="floating-label">Phone Number</label>
    <div className="phone-wrapper">
      <span className="phone-prefix">+62</span>
      <input
        type="tel"
        placeholder="812-3456-7890"
        value={FormData[type].phone}
        onChange={(e) => handleInputChange(e, type, "phone")}
        className="input-field phone-input"
      />
    </div>
  </div>
</div>

    </motion.div>
  );

  return (
    <motion.div
      className="form-container horizontal-form"
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      transition={{ duration: 0.5 }}
    >
    <div className="form-section">
  <h2 className="section-title">Select Status</h2>
  <div className="dropdowns">
    {/* Status Dropdown */}
    <div className="dropdown-group">
      <label className="dropdown-label">Administrative Division</label>
      <select
        value={status}
        onChange={(e) => {
          setStatus(e.target.value);
          setSelectedProvince("");
          setSelectedKabupaten("");
          setKabupatenList([]);
        }}
        className="select-status small-select"
      >
        <option value="">-- Select --</option>
        <option value="provincial">Provincial</option>
        <option value="kabupaten">Kabupaten</option>
      </select>
    </div>

    {/* Province Dropdown */}
    {status && (
      <div className="dropdown-group">
        <label className="dropdown-label">Province</label>
        <select
          value={selectedProvince}
          onChange={handleProvinceChange}
          className="select-status small-select"
        >
          <option value="">-- Select Province --</option>
          {provinces.map((prov) => (
            <option key={prov.adm_code} value={prov.LG}>
              {prov.LG}
            </option>
          ))}
        </select>
      </div>
    )}

    {/* Kabupaten Dropdown */}
    {status === "kabupaten" && (
      <div className="dropdown-group">
        <label className="dropdown-label">Kabupaten</label>
        <select
          value={selectedKabupaten}
          onChange={(e) => setSelectedKabupaten(e.target.value)}
          className="select-status small-select"
          disabled={!kabupatenList.length}
        >
          <option value="">-- Select Kabupaten --</option>
          {kabupatenList.map((kab) => (
            <option key={kab.adm_code} value={kab.LG}>
              {kab.LG}
            </option>
          ))}
        </select>
      </div>
    )}
  </div>
  {status && renderFormInputs(status)}
</div>

      <div className="form-section">
        <h3 className="section-title">Upload Excel Files</h3>
        

        <div className="file-upload-grid">
  {[
    "BridgeInventory",
    "CODE_AN_Parameters",
    "CODE_AN_UnitCostsPER",
    "CODE_AN_UnitCostsPERUnpaved",
    "CODE_AN_UnitCostsREH",
    "CODE_AN_UnitCostsRIGID",
    "CODE_AN_UnitCostsRM",
    "CODE_AN_UnitCostsUPGUnpaved",
    "CODE_AN_UnitCostsWidening",
    "CODE_AN_WidthStandards",
    "CulvertCondition",
    "CulvertInventory",
    "Link",
    "RetainingWallInventory",
    "RoadInventory",
    "TrafficVolume",
    "RoadCondition",
  ].map((key) => (
    <div key={key} className="file-upload-card">
      <div className="file-content">
        {!files[key] ? (
          <label className="full-area-label">
            <input
              type="file"
              accept=".xls,.xlsx"
              onChange={(e) => handleFileChange(e, key)}
              className="hidden-input"
            />
            <div className="upload-content">
              <span className="upload-icon">+</span>
              <span className="file-title">{key.split('_').join(' ')}</span>
              <p className="file-instruction">Click to upload</p>
            </div>
          </label>
        ) : (
          <div className="file-preview full-area-label">
            <div className="file-info">
              <span className="file-name">{files[key].name}</span>
              <button
                className="remove-btn"
                onClick={(e) => {
                  e.stopPropagation();
                  removeFile(key);
                }}
              >
                ×
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  ))}
</div>

        <button onClick={handleSubmit} className="button-submit">
          Generate Report
        </button>

        {output && (
          <div className="output-box">
            <pre>{JSON.stringify(output, null, 2)}</pre>
          </div>
        )}
      </div>
    </motion.div>
  );
}

export default Form;
