let workbookGlobal;

function initializeCatalog() {
  const savedData = localStorage.getItem('jsonData');
  
  if (savedData) {
    try {
      const jsonData = JSON.parse(savedData);
      if (validateJSONStructure(jsonData)) {
        renderCatalog(jsonData);
        return;
      } else {
        console.error("Formato de JSON guardado no es válido, cargando archivo por defecto.");
      }
    } catch (error) {
      console.error("Error al parsear JSON desde localStorage, cargando archivo por defecto.");
    }
  }
  
  loadJSONCatalog();
}

function validateJSONStructure(data) {
  if (!Array.isArray(data)) return false;
  return data.every(item => 
    item.hasOwnProperty('nombre') &&
    item.hasOwnProperty('precio') &&
    item.hasOwnProperty('descripcion')
  );
}

function loadJSONCatalog() {
  fetch('Datexce/catalogo.json')
    .then(response => {
      if (!response.ok) throw new Error('Error al cargar el archivo JSON');
      return response.json();
    })
    .then(data => {
      if (!validateJSONStructure(data)) {
        throw new Error('Formato de archivo JSON incorrecto');
      }
      localStorage.setItem('jsonData', JSON.stringify(data));
      renderCatalog(data);
    })
    .catch(error => {
      console.error('Error al procesar el archivo JSON:', error);
    });
}

function handleExcelLoad() {
  const savedData = localStorage.getItem('excelData');
  const uploadExcel = document.getElementById('uploadExcel');
  const sheetSelector = document.getElementById('sheetSelector');

  if (savedData) {
    workbookGlobal = JSON.parse(savedData);
    uploadExcel.style.display = 'none';
    sheetSelector.style.display = 'inline-block';

    sheetSelector.innerHTML = '<option value="">Selecciona un Producto</option>';
    workbookGlobal.SheetNames.forEach(function(sheetName, index) {
      const option = document.createElement('option');
      option.value = index;
      option.text = sheetName;
      sheetSelector.appendChild(option);
    });

    const selectedSheetIndex = localStorage.getItem('selectedSheetIndex');
    if (selectedSheetIndex) {
      sheetSelector.value = selectedSheetIndex;
      loadSheet();
    }
  } else {
    uploadExcel.style.display = 'inline-block';
    fetch('Datexce/Catálogo actualizado 05 de sep.xlsx')
      .then(response => {
        if (!response.ok) throw new Error('Error al cargar el archivo');
        return response.arrayBuffer();
      })
      .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        workbookGlobal = workbook;
        localStorage.setItem('excelData', JSON.stringify(workbookGlobal));

        sheetSelector.style.display = 'inline-block';
        sheetSelector.innerHTML = '<option value="">Selecciona un Producto</option>';
        workbook.SheetNames.forEach(function(sheetName, index) {
          const option = document.createElement('option');
          option.value = index;
          option.text = sheetName;
          sheetSelector.appendChild(option);
        });
      })
      .catch(error => {
        console.error('Error al cargar el archivo:', error);
      });
  }
}

document.getElementById('uploadExcel').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById('uploadExcel').style.display = 'none';

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    workbookGlobal = workbook;

    localStorage.setItem('excelData', JSON.stringify(workbookGlobal));

    const sheetSelector = document.getElementById('sheetSelector');
    sheetSelector.style.display = 'inline-block';
    sheetSelector.innerHTML = '<option value="">Selecciona un producto</option>';
    workbook.SheetNames.forEach(function (sheetName, index) {
      const option = document.createElement('option');
      option.value = index;
      option.text = sheetName;
      sheetSelector.appendChild(option);
    });
  };
  reader.readAsArrayBuffer(file);
}

function isBase64Image(base64String) {
  return typeof base64String === "string" && base64String.startsWith("data:image/");
}

function loadSheet() {
  const sheetIndex = document.getElementById('sheetSelector').value;

  if (sheetIndex === '') {
    document.getElementById('output').innerHTML = '';
    return;
  }

  localStorage.setItem('selectedSheetIndex', sheetIndex);

  const sheetName = workbookGlobal.SheetNames[sheetIndex];
  const sheet = workbookGlobal.Sheets[sheetName];
  const sheetRange = XLSX.utils.decode_range(sheet['!ref']);
  createCardsFromExcel(sheet, sheetRange);
}

function createCardsFromExcel(sheet, data) {
  const output = document.getElementById('output');
  output.innerHTML = ''; 
  let rowHtml = '<div class="row">';
  let cardCount = 0;

  for (let rowNum = data.s.r + 1; rowNum <= data.e.r; rowNum++) { 
    const productName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })]; 
    const productValue = productName ? productName.v : 'Sin nombre'; 

    rowHtml += `
      <div class="col-md-4 mt-3">
        <div class="card" style="width: 18rem;">
          <img src="https://via.placeholder.com/150" class="card-img-top" alt="Imagen de producto">
          <div class="card-body">
            <h5 class="card-title">${productValue}</h5>
            <p class="card-text" id="cardText${rowNum}">`;

    let pricesHtml = '<div style="line-height: 1.5;">'; 
    for (let colNum = 1; colNum <= data.e.c; colNum++) { 
      const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
      const cell = sheet[cellAddress];
      const cellValue = cell ? cell.v : '';

      if ([4, 6].includes(colNum) && !isNaN(cellValue)) { 
        const formattedPrice = parseFloat(cellValue).toLocaleString('es-CO', { style: 'currency', currency: 'COP' });
        pricesHtml += `<strong style="color: green;">${formattedPrice}</strong><br>`; 
      } else if (colNum === 5 && !isNaN(cellValue)) { 
        const ivaPercentage = (cellValue * 100).toFixed(0);
        pricesHtml += `<strong style="color: orange;">IVA ${ivaPercentage}%</strong><br>`; 
      } else if (isBase64Image(cellValue)) {
        rowHtml += `<img src="${cellValue}" width="100" class="mb-2"/>`;
      } else {
        rowHtml += `${cellValue} `;
      }
    }

    pricesHtml += '</div>'; 
    rowHtml += `${pricesHtml}</p>
            <span class="show-more" id="showMoreBtn${rowNum}">Leer más</span>
            <a href="https://wa.me/573163615434" class="btn btn-primary mt-2">Comunicarme</a>
          </div>
        </div>
      </div>`;

    cardCount++;
    
    if (cardCount % 3 === 0) { 
      rowHtml += '</div><div class="row">';
    }
  }

  rowHtml += '</div>';
  output.innerHTML = rowHtml;

  document.querySelectorAll('.show-more').forEach((button, index) => {
    button.addEventListener('click', function() {
      const cardText = document.getElementById(`cardText${index + 1}`);
      cardText.classList.toggle('expanded');
      button.textContent = cardText.classList.contains('expanded') ? 'Leer menos' : 'Leer más';
    });
  });
}

window.onload = function() {
  initializeCatalog();
  handleExcelLoad();
};
