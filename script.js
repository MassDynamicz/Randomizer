const numbers = [];
const usedNumbers = [];
const randomize = () => {
  const excludeNumbersInput = document.getElementById('excludeNumbers');
  const excludeNumbers = excludeNumbersInput.value.trim();
  const excludeNumbersArray = excludeNumbers === '' ? [] : excludeNumbers.split(',').map(n => parseInt(n, 10));
  const excludeRepeatsInput = document.getElementById('excludeRepeats');
  const excludeRepeats = excludeRepeatsInput.checked;

  if (excludeNumbersArray.some(isNaN)) {
    alert('Введите числа через запятую');
    return;
  }

  const filteredNumbers = numbers.filter(number => {
    const isExcluded = excludeNumbersArray.indexOf(number) !== -1;
    const isUsed = usedNumbers.indexOf(number) !== -1;
    return !isExcluded && (!excludeRepeats || !isUsed);
  });

  if (filteredNumbers.length === 0) {
    const noDataMessage = document.createElement('span');
    noDataMessage.textContent = 'Нет данных';
    document.getElementById('result').innerHTML = '';
    document.getElementById('result').appendChild(noDataMessage);
    return;
  }

  let randomNum;
  if (excludeRepeats) {
    const availableNumbers = filteredNumbers.filter(number => usedNumbers.indexOf(number) === -1);
    if (availableNumbers.length === 0) {
      usedNumbers.length = 0;
      randomNum = filteredNumbers[Math.floor(Math.random() * filteredNumbers.length)];
    } else {
      randomNum = availableNumbers[Math.floor(Math.random() * availableNumbers.length)];
    }
    usedNumbers.push(randomNum);
  } else {
    randomNum = filteredNumbers[Math.floor(Math.random() * filteredNumbers.length)];
  }
  animateNumber(randomNum);
};
// функция анимации
var animationInterval;
function animateNumber(num) {
  clearInterval(animationInterval);
  var currentNum = 0;
  var steps = 10;
  var diff = num - currentNum;
  var stepValue = Math.ceil(diff / steps);
  animationInterval = setInterval(function () {
    currentNum += stepValue;
    if (currentNum >= num) {
      clearInterval(animationInterval);
      currentNum = num;
    }
    document.getElementById('result').innerHTML = currentNum;
  }, 30);
}

// Загрузка файла Excel
var xhr = new XMLHttpRequest();
xhr.open('GET', 'data.xlsx', true);
xhr.responseType = 'arraybuffer';
xhr.onload = function (e) {
  if (xhr.status == 200) {
    var data = new Uint8Array(xhr.response);
    var workbook = XLSX.read(data, { type: 'array' });
    // Считывание данных из листа "numbers"
    var sheetName = 'numbers';
    var worksheet = workbook.Sheets[sheetName];
    var range = XLSX.utils.decode_range(worksheet['!ref']);

    for (var R = range.s.r; R <= range.e.r; ++R) {
      var number = worksheet[XLSX.utils.encode_cell({ r: R, c: 0 })].v;
      numbers.push(number);
    }
  }
};
xhr.send();
