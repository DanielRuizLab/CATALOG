let workbookGlobal;

function handleExcelLoad() {
  const fileName = 'Datexce/Catálogo prueba.xlsx';
  
  fetch(fileName)
    .then(response => { 
      if (!response.ok) throw new Error('El archivo Excel no está disponible o tiene un formato incorrecto');
      return response.arrayBuffer();
    })
    .then(data => {    
      const workbook = XLSX.read(data, { type: 'array' });
      console.log(`El archivo ${fileName} fue cargado correctamente`);

      workbookGlobal = workbook;

      const firstSheetName = workbook.SheetNames.find(name => name.toUpperCase() !== 'CATALOGO');
      if (firstSheetName) {
        const sheet = workbook.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(sheet['!ref']);
        createCardsFromExcel(sheet, range);
      }
    })
    .catch(error => {
      console.error('Error al cargar el archivo:', error);
    });
}
document.getElementById('uploadExcel').addEventListener('change', (event) => {
  const file = event.target.files[0];
  if (file) {
      console.log(`Archivo seleccionado: ${file.name}`);
  }
});

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  document.getElementById('uploadExcel').style.display = 'none';
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    workbookGlobal = workbook;

    const sheetSelector = document.getElementById('sheetSelector');
    sheetSelector.style.display = 'inline-block';
    sheetSelector.innerHTML = '<option value="">Selecciona un producto</option>';
    workbook.SheetNames.forEach(function (sheetName, index) {
      if (sheetName.toUpperCase() === 'CATALOGO') return;
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

    const imageName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 7 })]; 
    const imageUrl = imageName ? `img/imgoppo/${imageName.v}` : 'https://via.placeholder.com/150';  

    rowHtml += `  
      <div class="col-12 col-sm-6 col-md-4 mb-4"> 
        <div class="card" style="width: 100%;">
          <img src="${imageUrl}" class="card-img-top" alt="Imagen de producto">
          <div class="card-body"> 
            <h5 class="card-title">${productValue}</h5>
            <p class="card-text" id="cardText${rowNum}">
    `; 

    let pricesHtml = '<div style="line-height: 1.5;">';
    let unitPrice = null;
    let iva = null;

    for (let colNum = 1; colNum <= data.e.c; colNum++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
      const cell = sheet[cellAddress];
      const cellValue = cell ? cell.v : '';

      if (colNum === 4 && !isNaN(cellValue)) { 
        unitPrice = parseFloat(cellValue);
      } else if (colNum === 5 && !isNaN(cellValue)) { 
        iva = parseFloat(cellValue);
      } else if (colNum !== 4 && colNum !== 5) { 
        if (isBase64Image(cellValue)) {
          rowHtml += `<img src="${cellValue}" width="100" class="mb-2"/>`;
        } else {
          // Eliminar texto que contenga nombres de archivos de imagen (.jpg, .png, etc.)
          const cleanText = typeof cellValue === 'string'
            ? cellValue.replace(/\s*\/?\s*[\w\s-]+\.(jpg|jpeg|png|webp)/gi, '') 
            : cellValue;
          rowHtml += `${cleanText} `;
        }
      }
    }

    if (unitPrice !== null && iva !== null) {
      const finalPrice = unitPrice + (unitPrice * iva);
      const formattedFinalPrice = Math.round(finalPrice).toLocaleString('es-CO');
      pricesHtml += `<strong style="color: green;">$${formattedFinalPrice}</strong><br>`;
    }

    pricesHtml += '</div>';
    rowHtml += `${pricesHtml}</p>
      <span class="show-more" id="showMoreBtn${rowNum}">Leer más</span>
      <a href="https://wa.me/573163615434" class="btn btn-primary mt-2 btn-whatsapp" target="_blank">Hablar con un Asesor</a>  
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
  alert("⚠️😁 Los productos están sujetos a cambio de referencias y precios.⚠️");
  handleExcelLoad();
};