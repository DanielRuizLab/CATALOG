let workbookGlobal;

function handleExcelLoad() {
  const uploadExcel = document.getElementById('uploadExcel');
  const sheetSelector = document.getElementById('sheetSelector');

  uploadExcel.style.display = 'inline-block';
  
  fetch('Datexce/Catálogo actualizado 12 de nov.xlsx')
    .then(response => {
      if (!response.ok) throw new Error('Error al cargar el archivo');
      return response.arrayBuffer();
    })
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      workbookGlobal = workbook;

      sheetSelector.style.display = 'inline-block';
      sheetSelector.innerHTML = '<option value="">Selecciona una Categoria</option>';
      workbook.SheetNames.forEach(function (sheetName, index) {
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
    const imageUrl = imageName ? `img/${imageName.v}` : 'https://via.placeholder.com/150'; 

    rowHtml += `  
      <div class="col-lg-4 col-md-6 col-sm-12 mt-3">
        <div class="card" style="width: 18rem;">
          <img src="${imageUrl}" class="card-img-top" alt="Imagen de producto">
          <div class="card-body">
            <h5 class="card-title">${productValue}</h5>
            <p class="card-text" id="cardText${rowNum}">
    `;

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
      <a href="https://wa.me/573163615434" class="btn btn-primary mt-2" target="_blank">Comunicarme</a>
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
}window.onload = function() {
  handleExcelLoad();
};
