/**
 * Renaissance Global Limited — Summary Sheet Creation Logic
 * Form handling for:
 * 1. Customer
 * 2. File: Pricing Sheet (Drag-and-Drop / Browse)
 * 3. Diamond Quality
 */

document.addEventListener('DOMContentLoaded', () => {
  // --- Auth Session Verification ---
  const authSession = (() => {
    try {
      const raw = sessionStorage.getItem('rg_auth_session') || localStorage.getItem('rg_auth_session');
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  })();

  if (!authSession || !authSession.authenticated) {
    window.location.replace('index.html');
    return;
  }

  // Populate logged in user badge
  const headerUserEmail = document.getElementById('header-user-email');
  if (headerUserEmail && authSession.email) {
    headerUserEmail.textContent = authSession.email;
  }

  // Bind Sign Out action
  const headerLogoutBtn = document.getElementById('header-logout-btn');
  if (headerLogoutBtn) {
    headerLogoutBtn.addEventListener('click', (e) => {
      e.preventDefault();
      sessionStorage.removeItem('rg_auth_session');
      localStorage.removeItem('rg_auth_session');
      window.location.replace('index.html');
    });
  }

  // Form & Input elements
  const form = document.getElementById('summary-sheet-form');
  const inputCustomer = document.getElementById('input-customer');
  const inputCollection = document.getElementById('input-collection');
  const inputPricingSheet = document.getElementById('input-pricing-sheet');
  const inputDiamondQuality = document.getElementById('input-diamond-quality');
  const customerError = document.getElementById('customer-error');
  const fileError = document.getElementById('file-error');
  const qualityError = document.getElementById('quality-error');

  // Pricing Basis and Collection chips
  const chipCollectionBtns = document.querySelectorAll('.chip-collection');
  const priceBasisInputs = document.querySelectorAll('input[name="price_basis"]');
  const priceBasisOptions = document.querySelectorAll('.price-basis-option');

  // File Upload Drop Zone elements
  const fileDropZone = document.getElementById('file-drop-zone');
  const dropZoneIdle = document.getElementById('drop-zone-idle');
  const fileSelectedCard = document.getElementById('file-selected-card');
  const displayFileName = document.getElementById('display-file-name');
  const displayFileSize = document.getElementById('display-file-size');
  const fileRemoveBtn = document.getElementById('file-remove-btn');
  const browseFileBtn = document.getElementById('browse-file-btn');

  // Buttons & Controls
  const btnCreateSummary = document.getElementById('btn-create-summary');
  const btnSpinner = document.getElementById('btn-spinner');
  const btnResetForm = document.getElementById('btn-reset-form');
  const chipBtns = document.querySelectorAll('.chip-btn:not(.chip-collection)');
  const qualityPills = document.querySelectorAll('.quality-pill');

  // Preview & Result Elements
  const pipelineStatusBadge = document.getElementById('pipeline-status-badge');
  const progressContainer = document.getElementById('progress-container');
  const progressStepText = document.getElementById('progress-step-text');
  const progressPercentText = document.getElementById('progress-percent-text');
  const progressBarFill = document.getElementById('progress-bar-fill');
  const previewIdleView = document.getElementById('preview-idle-view');
  const previewResultView = document.getElementById('preview-result-view');
  const resultTimestamp = document.getElementById('result-timestamp');
  const resCustomer = document.getElementById('res-customer');
  const resCollection = document.getElementById('res-collection');
  const resPriceBasis = document.getElementById('res-price-basis');
  const resSkus = document.getElementById('res-skus');
  const genFilename = document.getElementById('gen-filename');
  const btnDownloadSheet = document.getElementById('btn-download-sheet');
  const toastContainer = document.getElementById('toast-container');

  // Stored state
  let selectedFile = null;

  // --- Quick Select Customer Chips ---
  chipBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const cust = btn.getAttribute('data-customer');
      if (cust && inputCustomer) {
        inputCustomer.value = cust;
        inputCustomer.classList.remove('invalid');
        customerError.style.display = 'none';
        inputCustomer.focus();
      }
    });
  });

  // --- Quick Select Collection Chips ---
  chipCollectionBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const col = btn.getAttribute('data-collection');
      if (col && inputCollection) {
        inputCollection.value = col;
        inputCollection.focus();
      }
    });
  });

  // --- Pricing Basis Radio Card Change Handlers ---
  priceBasisInputs.forEach(radio => {
    radio.addEventListener('change', () => {
      priceBasisOptions.forEach(opt => opt.classList.remove('active'));
      const parentLabel = radio.closest('.price-basis-option');
      if (parentLabel) {
        parentLabel.classList.add('active');
      }
    });
  });

  // --- Quick Select Diamond Quality Pills ---
  qualityPills.forEach(pill => {
    pill.addEventListener('click', () => {
      const qual = pill.getAttribute('data-quality');
      inputDiamondQuality.value = qual;
      inputDiamondQuality.classList.remove('invalid');
      qualityError.style.display = 'none';

      qualityPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
    });
  });

  // Sync select change with pills
  inputDiamondQuality.addEventListener('change', () => {
    const currentVal = inputDiamondQuality.value;
    qualityPills.forEach(p => {
      if (p.getAttribute('data-quality') === currentVal) {
        p.classList.add('active');
      } else {
        p.classList.remove('active');
      }
    });
    inputDiamondQuality.classList.remove('invalid');
    qualityError.style.display = 'none';
  });

  // --- File Drag and Drop Handlers ---
  if (browseFileBtn) {
    browseFileBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      inputPricingSheet.click();
    });
  }

  fileDropZone.addEventListener('click', () => {
    if (!selectedFile) {
      inputPricingSheet.click();
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    fileDropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileDropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    fileDropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileDropZone.classList.remove('drag-over');
    });
  });

  fileDropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
      handleFileSelection(files[0]);
    }
  });

  inputPricingSheet.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFileSelection(e.target.files[0]);
    }
  });

  function handleFileSelection(file) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExtensions.some(ext => fileName.endsWith(ext));

    if (!isValid) {
      showToast('Please upload a valid Excel or CSV spreadsheet (.xlsx, .xls, .csv)', 'tan');
      return;
    }

    selectedFile = file;
    displayFileName.textContent = file.name;
    displayFileSize.textContent = formatBytes(file.size) + ' • Spreadsheet Document';

    dropZoneIdle.style.display = 'none';
    fileSelectedCard.style.display = 'flex';
    fileDropZone.classList.add('has-file');

    fileDropZone.classList.remove('invalid');
    fileError.style.display = 'none';

    showToast(`Pricing Sheet "${file.name}" attached.`, 'green');
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  if (fileRemoveBtn) {
    fileRemoveBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      selectedFile = null;
      inputPricingSheet.value = '';
      dropZoneIdle.style.display = 'flex';
      fileSelectedCard.style.display = 'none';
      fileDropZone.classList.remove('has-file');
      showToast('File removed', 'tan');
    });
  }

  // --- Input Clearing on Typing ---
  inputCustomer.addEventListener('input', () => {
    if (inputCustomer.value.trim().length > 0) {
      inputCustomer.classList.remove('invalid');
      customerError.style.display = 'none';
    }
  });

  // --- Form Reset Handler ---
  btnResetForm.addEventListener('click', () => {
    form.reset();
    selectedFile = null;
    dropZoneIdle.style.display = 'flex';
    fileSelectedCard.style.display = 'none';
    fileDropZone.classList.remove('has-file');
    qualityPills.forEach(p => p.classList.remove('active'));

    inputCustomer.classList.remove('invalid');
    fileDropZone.classList.remove('invalid');
    inputDiamondQuality.classList.remove('invalid');
    customerError.style.display = 'none';
    fileError.style.display = 'none';
    qualityError.style.display = 'none';

    // Reset preview
    progressContainer.style.display = 'none';
    previewResultView.style.display = 'none';
    previewIdleView.style.display = 'block';
    pipelineStatusBadge.textContent = 'Standby';
    pipelineStatusBadge.className = 'status-pill-badge';

    showToast('Form reset to default state.', 'tan');
  });

  // --- Form Submission & Generation Simulation ---
  form.addEventListener('submit', (e) => {
    e.preventDefault();

    let hasError = false;

    // 1. File is strictly required
    if (!selectedFile) {
      fileDropZone.classList.add('invalid');
      fileError.style.display = 'block';
      showToast('Please upload a Pricing Sheet spreadsheet.', 'tan');
      return;
    } else {
      fileDropZone.classList.remove('invalid');
      fileError.style.display = 'none';
    }

    // 2. Customer, Collection, Pricing Basis & Diamond Quality
    const customerVal = inputCustomer.value.trim() || 'CLIENT RECAP';
    const collectionVal = inputCollection ? inputCollection.value.trim() : '';
    const selectedBasisRadio = document.querySelector('input[name="price_basis"]:checked');
    const priceBasisVal = selectedBasisRadio ? selectedBasisRadio.value : 'MEMO_COST';
    const qualityVal = inputDiamondQuality.value || 'From Pricing Sheet';

    // Begin Generation Sequence
    btnCreateSummary.disabled = true;
    btnSpinner.style.display = 'inline-block';
    pipelineStatusBadge.textContent = 'Processing';
    pipelineStatusBadge.className = 'status-pill-badge badge-processing';

    previewIdleView.style.display = 'none';
    previewResultView.style.display = 'none';
    progressContainer.style.display = 'block';

    const steps = [
      { text: 'Uploading pricing sheet to backend engine...', pct: 20 },
      { text: 'Applying server master Recap Template & Code Mappings...', pct: 50 },
      { text: 'Resolving style items, Disney titles & diamond fractions...', pct: 75 },
      { text: 'Generating Client Recap workbook...', pct: 90 }
    ];

    let currentStep = 0;
    const progressTimer = setInterval(() => {
      if (currentStep < steps.length) {
        progressStepText.textContent = steps[currentStep].text;
        progressPercentText.textContent = steps[currentStep].pct + '%';
        progressBarFill.style.width = steps[currentStep].pct + '%';
        currentStep++;
      }
    }, 450);

    // Prepare Multipart Form Data for Backend REST API
    const formData = new FormData();
    formData.append('customer', customerVal);
    formData.append('collection_name', collectionVal);
    formData.append('price_basis', priceBasisVal);
    formData.append('pricing_sheet', selectedFile);
    formData.append('diamond_quality', qualityVal);

    // Call Backend REST API
    fetch('/api/summary-sheet', {
      method: 'POST',
      body: formData
    })
    .then(async (response) => {
      clearInterval(progressTimer);
      const result = await response.json();
      if (!response.ok || !result.success) {
        throw new Error(result.error || 'Server error processing pricing sheet');
      }
      return result.data;
    })
    .then((data) => {
      progressBarFill.style.width = '100%';
      progressPercentText.textContent = '100%';
      progressStepText.textContent = 'Completed!';

      setTimeout(() => {
        btnCreateSummary.disabled = false;
        btnSpinner.style.display = 'none';
        progressContainer.style.display = 'none';

        // Update result card with actual backend parsed data
        pipelineStatusBadge.textContent = 'Completed';
        pipelineStatusBadge.className = 'status-pill-badge badge-completed';

        resultTimestamp.textContent = `Processed on server at ${data.processed_at}`;
        if (resCustomer) resCustomer.textContent = data.customer;
        if (resCollection) resCollection.textContent = data.collection_name || 'Standard';
        if (resPriceBasis) resPriceBasis.textContent = data.price_basis === 'SELLING_PRICE' ? 'Selling Price' : 'MEMO COST';

        if (resSkus) {
          resSkus.textContent = `${data.total_rows_parsed} Line Items`;
        }

        genFilename.textContent = data.generated_filename;

        // Set download action to direct backend download endpoint
        btnDownloadSheet.onclick = () => {
          window.location.href = data.download_url;
          showToast(`Downloading "${data.generated_filename}" from server...`, 'green');
        };

        previewResultView.style.display = 'block';

        showToast(`Client Recap generated by backend for ${data.customer}!`, 'green');
      }, 350);
    })
    .catch((err) => {
      clearInterval(progressTimer);
      btnCreateSummary.disabled = false;
      btnSpinner.style.display = 'none';
      progressContainer.style.display = 'none';
      previewIdleView.style.display = 'block';
      pipelineStatusBadge.textContent = 'Standby';
      pipelineStatusBadge.className = 'status-pill-badge';

      showToast(`Error: ${err.message}`, 'tan');
      console.error('Summary Sheet API Error:', err);
    });
  });

  // --- Toast Notification Helper ---
  function showToast(message, type = 'green') {
    const toast = document.createElement('div');
    toast.className = `toast ${type === 'tan' ? 'toast-tan' : ''}`;
    toast.innerHTML = `
      <svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
        <polyline points="22 4 12 14.01 9 11.01"></polyline>
      </svg>
      <span>${message}</span>
    `;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(30px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }
});
