// czCalc — Calculator Logic
'use strict';

let expression = '';
let currentNum = '0';
let operator = null;
let prevNum = null;
let justCalculated = false;

const elExpr = document.getElementById('expression');
const elResult = document.getElementById('result');

function updateDisplay() {
  elResult.textContent = formatNumber(currentNum);
  elResult.classList.remove('flash');
  void elResult.offsetWidth; // reflow
  elResult.classList.add('flash');
}

function formatNumber(val) {
  if (val === 'Error') return val;
  const num = parseFloat(val);
  if (isNaN(num)) return '0';
  // Limit display to 12 significant digits
  if (Math.abs(num) >= 1e12) return num.toExponential(6);
  const str = parseFloat(num.toPrecision(12)).toString();
  // Add commas for integer part
  const parts = str.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function appendNum(n) {
  if (justCalculated) {
    currentNum = '0';
    expression = '';
    justCalculated = false;
  }
  if (currentNum === '0' && n !== '.') {
    currentNum = n;
  } else if (currentNum.length < 15) {
    currentNum += n;
  }
  updateDisplay();
  clearActiveOp();
}

function appendDot() {
  if (justCalculated) {
    currentNum = '0';
    expression = '';
    justCalculated = false;
  }
  if (!currentNum.includes('.')) {
    currentNum += '.';
    updateDisplay();
  }
}

function appendOp(op) {
  justCalculated = false;
  if (operator && prevNum !== null) {
    const result = compute(parseFloat(prevNum), parseFloat(currentNum), operator);
    currentNum = String(result);
    updateDisplay();
  }
  prevNum = currentNum;
  operator = op;
  expression = formatNumber(prevNum) + ' ' + opSymbol(op);
  elExpr.textContent = expression;
  currentNum = '0';
  setActiveOp(op);
}

function calculate() {
  if (operator && prevNum !== null) {
    expression = formatNumber(prevNum) + ' ' + opSymbol(operator) + ' ' + formatNumber(currentNum) + ' =';
    elExpr.textContent = expression;
    const result = compute(parseFloat(prevNum), parseFloat(currentNum), operator);
    currentNum = String(result);
    updateDisplay();
    operator = null;
    prevNum = null;
    justCalculated = true;
    clearActiveOp();
  }
}

function compute(a, b, op) {
  switch (op) {
    case '+': return a + b;
    case '-': return a - b;
    case '*': return a * b;
    case '/': return b === 0 ? 'Error' : a / b;
    default:  return b;
  }
}

function clearAll() {
  expression = '';
  currentNum = '0';
  operator = null;
  prevNum = null;
  justCalculated = false;
  elExpr.textContent = '';
  updateDisplay();
  clearActiveOp();
}

function toggleSign() {
  if (currentNum !== '0' && currentNum !== 'Error') {
    currentNum = String(-parseFloat(currentNum));
    updateDisplay();
  }
}

function percentage() {
  if (currentNum !== 'Error') {
    currentNum = String(parseFloat(currentNum) / 100);
    updateDisplay();
  }
}

function opSymbol(op) {
  switch (op) {
    case '+': return '+';
    case '-': return '\u2212';
    case '*': return '\u00D7';
    case '/': return '\u00F7';
    default:  return op;
  }
}

function setActiveOp(op) {
  clearActiveOp();
  const symbols = { '/': '\u00F7', '*': '\u00D7', '-': '\u2212', '+': '+' };
  document.querySelectorAll('.btn-op').forEach(btn => {
    if (btn.textContent === symbols[op]) btn.classList.add('active');
  });
}

function clearActiveOp() {
  document.querySelectorAll('.btn-op').forEach(btn => btn.classList.remove('active'));
}

// Keyboard support
document.addEventListener('keydown', e => {
  if (e.key >= '0' && e.key <= '9') appendNum(e.key);
  else if (e.key === '.') appendDot();
  else if (e.key === '+') appendOp('+');
  else if (e.key === '-') appendOp('-');
  else if (e.key === '*') appendOp('*');
  else if (e.key === '/') { e.preventDefault(); appendOp('/'); }
  else if (e.key === 'Enter' || e.key === '=') calculate();
  else if (e.key === 'Escape') clearAll();
  else if (e.key === 'Backspace') {
    if (currentNum.length > 1) {
      currentNum = currentNum.slice(0, -1);
    } else {
      currentNum = '0';
    }
    updateDisplay();
  }
});
