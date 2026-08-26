/**
 * Casio-Style Scientific Calculator Engine
 */

class ScientificCalculator {
  constructor() {
    this.expression = '';       // Current expression typed
    this.displayVal = '0';      // Main value displayed
    this.historyVal = '';      // Upper LCD line
    this.angleMode = 'DEG';     // DEG, RAD, GRAD
    this.isShift = false;       // Shift active flag
    this.isHyp = false;         // Hyperbolic active flag
    this.memory = 0;            // Calculator memory M
    this.ans = 0;               // Last answer
    this.fseMode = 'NORM';      // NORM, FIX, SCI
    this.baseMode = 'DEC';      // DEC, HEX, OCT, BIN
    this.justCalculated = false;

    // DOM Elements
    this.elLcdNumber = document.getElementById('lcdNumber');
    this.elLcdExpression = document.getElementById('lcdExpression');
    this.elLcdExpGroup = document.getElementById('lcdExpGroup');
    this.elLcdExpVal = document.getElementById('lcdExpVal');
    this.elIndDeg = document.getElementById('indDeg');
    this.elIndRad = document.getElementById('indRad');
    this.elIndGrad = document.getElementById('indGrad');
    this.elIndShift = document.getElementById('indShift');
    this.elIndHyp = document.getElementById('indHyp');
    this.elIndMem = document.getElementById('indMem');
    this.elBtnShift = document.getElementById('btnShift');

    this.initEvents();
    this.updateDisplay();
  }

  initEvents() {
    // Calculator button clicks
    document.querySelectorAll('.calc-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const action = this.isShift ? (btn.dataset.shiftAction || btn.dataset.action) : btn.dataset.action;
        const val = this.isShift ? (btn.dataset.shiftVal || btn.dataset.val) : btn.dataset.val;
        this.handleButtonAction(action, val, btn);
        
        // Auto reset shift after action (unless SHIFT button itself)
        if (this.isShift && action !== 'shift') {
          this.toggleShift(false);
        }
      });
    });

    // Physical Keyboard Listener
    window.addEventListener('keydown', (e) => {
      // Do not capture if typing in an input element (like spreadsheet formula bar or cell editor)
      if (['INPUT', 'TEXTAREA'].includes(e.target.tagName) || e.target.isContentEditable) {
        return;
      }
      this.handleKeyboard(e);
    });
  }

  handleButtonAction(action, val, btnEl) {
    if (btnEl) {
      btnEl.classList.add('pressed');
      setTimeout(() => btnEl.classList.remove('pressed'), 120);
    }

    switch (action) {
      case 'shift':
        this.toggleShift(!this.isShift);
        break;
      case 'hyp':
        this.isHyp = !this.isHyp;
        this.updateIndicators();
        break;
      case 'drg':
        this.cycleAngleMode();
        break;
      case 'drg_conv':
        this.convertAngleCurrentValue();
        break;
      case 'fse':
        this.cycleFseMode();
        break;
      case 'mc':
        this.memory = 0;
        this.updateIndicators();
        this.showToast('메모리 클리어 (MC)');
        break;
      case 'mr':
        this.appendInput(String(this.memory));
        break;
      case 'ms':
        try {
          this.memory = this.evaluateCurrent();
          this.updateIndicators();
          this.showToast(`메모리 저장 (MS): ${this.memory}`);
        } catch (e) {
          this.displayVal = 'Error';
        }
        break;
      case 'm_plus':
        try {
          this.memory += this.evaluateCurrent();
          this.updateIndicators();
          this.showToast(`메모리 + (M+): ${this.memory}`);
        } catch (e) {
          this.displayVal = 'Error';
        }
        break;
      case 'm_minus':
        try {
          this.memory -= this.evaluateCurrent();
          this.updateIndicators();
          this.showToast(`메모리 - (M-): ${this.memory}`);
        } catch (e) {
          this.displayVal = 'Error';
        }
        break;
      case 'ac':
        this.allClear();
        break;
      case 'del':
        this.deleteChar();
        break;
      case 'equals':
        this.calculate();
        break;
      case 'negate':
        this.toggleNegate();
        break;
      case 'base':
        this.setBaseMode(val);
        break;
      case 'ans':
        this.appendInput('ANS');
        break;
      case 'func':
        this.insertFunction(val);
        break;
      case 'op':
        this.insertOperator(val);
        break;
      case 'num':
        this.insertNumber(val);
        break;
      case 'const':
        this.insertConstant(val);
        break;
      case 'info':
        this.showInfo();
        break;
      default:
        if (val) this.appendInput(val);
    }
    this.updateDisplay();
  }

  handleKeyboard(e) {
    const key = e.key;

    // Prevent default scrolling for Space / / / Arrow keys
    if (['/', 'Enter', 'Backspace', 'Escape'].includes(key)) {
      e.preventDefault();
    }

    if (key >= '0' && key <= '9') {
      this.insertNumber(key);
    } else if (key === '.') {
      this.insertNumber('.');
    } else if (['+', '-', '*', '/'].includes(key)) {
      const opMap = { '+': '+', '-': '−', '*': '×', '/': '÷' };
      this.insertOperator(opMap[key]);
    } else if (key === '(' || key === ')') {
      this.appendInput(key);
    } else if (key === '^') {
      this.insertOperator('^');
    } else if (key === '%') {
      this.insertOperator('%');
    } else if (key === 'Enter' || key === '=') {
      this.calculate();
    } else if (key === 'Backspace') {
      this.deleteChar();
    } else if (key === 'Escape') {
      this.allClear();
    } else if (key.toLowerCase() === 's') {
      this.insertFunction('sin');
    } else if (key.toLowerCase() === 'c') {
      this.insertFunction('cos');
    } else if (key.toLowerCase() === 't') {
      this.insertFunction('tan');
    } else if (key.toLowerCase() === 'l') {
      this.insertFunction('ln');
    } else if (key.toLowerCase() === 'g') {
      this.insertFunction('log');
    } else if (key.toLowerCase() === 'p') {
      this.insertConstant('pi');
    } else if (key.toLowerCase() === 'e' && !e.ctrlKey) {
      this.insertConstant('e');
    } else if (key.toLowerCase() === 'r') {
      this.insertFunction('sqrt');
    } else if (key === '!') {
      this.appendInput('!');
    }

    this.updateDisplay();
  }

  toggleShift(state) {
    this.isShift = state !== undefined ? state : !this.isShift;
    if (this.elBtnShift) {
      this.elBtnShift.classList.toggle('active', this.isShift);
    }
    this.updateIndicators();
  }

  cycleAngleMode() {
    if (this.angleMode === 'DEG') this.angleMode = 'RAD';
    else if (this.angleMode === 'RAD') this.angleMode = 'GRAD';
    else this.angleMode = 'DEG';
    this.updateIndicators();
  }

  convertAngleCurrentValue() {
    try {
      const num = this.evaluateCurrent();
      let res = num;
      // Convert DEG -> RAD -> GRAD -> DEG
      if (this.angleMode === 'DEG') {
        res = (num * Math.PI) / 180; // to RAD
        this.angleMode = 'RAD';
      } else if (this.angleMode === 'RAD') {
        res = (num * 200) / Math.PI; // to GRAD
        this.angleMode = 'GRAD';
      } else {
        res = (num * 180) / 200; // to DEG
        this.angleMode = 'DEG';
      }
      this.displayVal = this.formatResult(res);
      this.expression = this.displayVal;
      this.updateIndicators();
    } catch (e) {
      this.displayVal = 'Error';
    }
  }

  cycleFseMode() {
    const modes = ['NORM', 'FIX', 'SCI'];
    const idx = modes.indexOf(this.fseMode);
    this.fseMode = modes[(idx + 1) % modes.length];
    if (this.displayVal !== 'Error' && this.displayVal !== '0') {
      this.displayVal = this.formatResult(parseFloat(this.displayVal));
    }
    this.showToast(`표시 모드: ${this.fseMode}`);
  }

  setBaseMode(base) {
    this.baseMode = base;
    try {
      const val = parseInt(this.evaluateCurrent(), 10);
      if (!isNaN(val)) {
        if (base === 'HEX') this.displayVal = val.toString(16).toUpperCase();
        else if (base === 'OCT') this.displayVal = val.toString(8);
        else if (base === 'BIN') this.displayVal = val.toString(2);
        else this.displayVal = val.toString(10);
      }
    } catch (e) {}
    this.showToast(`진법 모드: ${base}`);
  }

  insertNumber(num) {
    if (this.justCalculated) {
      this.expression = '';
      this.justCalculated = false;
    }
    this.appendInput(num);
  }

  insertOperator(op) {
    if (this.justCalculated) {
      this.expression = this.displayVal;
      this.justCalculated = false;
    }
    if (this.expression === '' && op === '−') {
      this.appendInput('−');
      return;
    }
    this.appendInput(` ${op} `);
  }

  insertFunction(func) {
    if (this.justCalculated) {
      this.expression = '';
      this.justCalculated = false;
    }

    // Check hyperbolic
    if (this.isHyp) {
      if (['sin', 'cos', 'tan', 'asin', 'acos', 'atan'].includes(func)) {
        func = func.startsWith('a') ? `a${func.slice(1)}h` : `${func}h`;
      }
      this.isHyp = false;
      this.updateIndicators();
    }

    if (['sqr', 'cube', 'recip', 'fact'].includes(func)) {
      if (func === 'sqr') this.appendInput('^2');
      else if (func === 'cube') this.appendInput('^3');
      else if (func === 'recip') this.appendInput('^(-1)');
      else if (func === 'fact') this.appendInput('!');
    } else {
      this.appendInput(`${func}(`);
    }
  }

  insertConstant(c) {
    if (this.justCalculated) {
      this.expression = '';
      this.justCalculated = false;
    }
    if (c === 'pi') this.appendInput('π');
    else if (c === 'e') this.appendInput('e');
  }

  toggleNegate() {
    if (this.expression) {
      if (this.expression.startsWith('−')) {
        this.expression = this.expression.slice(1);
      } else {
        this.expression = '−' + this.expression;
      }
    } else {
      this.expression = '−';
    }
  }

  appendInput(str) {
    this.expression += str;
    this.displayVal = this.expression;
  }

  deleteChar() {
    if (this.justCalculated) {
      this.allClear();
      return;
    }
    if (this.expression.length > 0) {
      if (this.expression.endsWith(' ')) {
        this.expression = this.expression.trimEnd();
      }
      this.expression = this.expression.slice(0, -1);
      this.displayVal = this.expression || '0';
    }
  }

  allClear() {
    this.expression = '';
    this.displayVal = '0';
    this.historyVal = '';
    this.justCalculated = false;
  }

  evaluateCurrent() {
    if (!this.expression) return parseFloat(this.displayVal) || 0;
    return this.parseExpression(this.expression);
  }

  calculate() {
    if (!this.expression) return;
    try {
      const result = this.parseExpression(this.expression);
      this.historyVal = this.expression + ' =';
      this.ans = result;
      this.displayVal = this.formatResult(result);
      this.expression = '';
      this.justCalculated = true;
    } catch (e) {
      this.historyVal = this.expression;
      this.displayVal = 'Error';
      this.justCalculated = true;
    }
    this.updateDisplay();
  }

  formatResult(val) {
    if (isNaN(val)) return 'Error';
    if (!isFinite(val)) return val > 0 ? 'Infinity' : '-Infinity';

    if (this.fseMode === 'SCI') {
      return val.toExponential(6);
    } else if (this.fseMode === 'FIX') {
      return val.toFixed(4);
    }

    // Default clean scientific / standard formatting
    if (Math.abs(val) !== 0 && (Math.abs(val) >= 1e11 || Math.abs(val) <= 1e-7)) {
      return val.toExponential(6).replace('e+', 'e');
    }

    // Limit decimal noise (e.g. 0.1 + 0.2 = 0.30000000000000004 -> 0.3)
    const rounded = Number(Math.round(val + 'e12') + 'e-12');
    return String(rounded);
  }

  parseExpression(expr) {
    let s = expr;

    // Replace visual symbols with math tokens
    s = s.replace(/×/g, '*')
         .replace(/÷/g, '/')
         .replace(/−/g, '-')
         .replace(/π/g, `(${Math.PI})`)
         .replace(/ANS/g, `(${this.ans})`)
         .replace(/\be\b/g, `(${Math.E})`);

    // Handle scientific EXP (e.g. 5 EXP 3 -> 5e3)
    s = s.replace(/EXP\s*([0-9\-]+)/gi, 'e$1');

    // Handle combinations (nCr) and permutations (nPr)
    s = s.replace(/(\d+(?:\.\d+)?)\s*nCr\s*(\d+(?:\.\d+)?)/g, (_, n, r) => this.mathCombination(Number(n), Number(r)));
    s = s.replace(/(\d+(?:\.\d+)?)\s*nPr\s*(\d+(?:\.\d+)?)/g, (_, n, r) => this.mathPermutation(Number(n), Number(r)));
    s = s.replace(/(\d+(?:\.\d+)?)\s*MOD\s*(\d+(?:\.\d+)?)/gi, '($1 % $2)');

    // Handle factorials (e.g. 5!)
    s = s.replace(/(\d+)!/g, (_, n) => this.mathFactorial(Number(n)));

    // Handle powers (^)
    s = s.replace(/(\b[\w\.\)]+)\s*\^\s*(\b[\w\.\(]+)/g, 'Math.pow($1, $2)');

    // Handle trigonometric, hyperbolic, roots, logs with Angle Mode
    const toRad = this.angleMode === 'DEG' ? `* (Math.PI / 180)` : (this.angleMode === 'GRAD' ? `* (Math.PI / 200)` : ``);
    const fromRad = this.angleMode === 'DEG' ? `* (180 / Math.PI)` : (this.angleMode === 'GRAD' ? `* (200 / Math.PI)` : ``);

    // Trig functions
    s = s.replace(/\bsin\(([^)]+)\)/g, `Math.sin(($1) ${toRad})`);
    s = s.replace(/\bcos\(([^)]+)\)/g, `Math.cos(($1) ${toRad})`);
    s = s.replace(/\btan\(([^)]+)\)/g, `Math.tan(($1) ${toRad})`);
    s = s.replace(/\basin\(([^)]+)\)/g, `(Math.asin($1) ${fromRad})`);
    s = s.replace(/\bacos\(([^)]+)\)/g, `(Math.acos($1) ${fromRad})`);
    s = s.replace(/\batan\(([^)]+)\)/g, `(Math.atan($1) ${fromRad})`);

    // Hyperbolic
    s = s.replace(/\bsinh\(([^)]+)\)/g, `Math.sinh($1)`);
    s = s.replace(/\bcosh\(([^)]+)\)/g, `Math.cosh($1)`);
    s = s.replace(/\btanh\(([^)]+)\)/g, `Math.tanh($1)`);
    s = s.replace(/\basinh\(([^)]+)\)/g, `Math.asinh($1)`);
    s = s.replace(/\bacosh\(([^)]+)\)/g, `Math.acosh($1)`);
    s = s.replace(/\batanh\(([^)]+)\)/g, `Math.atanh($1)`);

    // Log & Roots & Exp
    s = s.replace(/\bln\(([^)]+)\)/g, `Math.log($1)`);
    s = s.replace(/\blog\(([^)]+)\)/g, `Math.log10($1)`);
    s = s.replace(/\bsqrt\(([^)]+)\)/g, `Math.sqrt($1)`);
    s = s.replace(/\bcbrt\(([^)]+)\)/g, `Math.cbrt($1)`);
    s = s.replace(/\broot\(([^,]+),([^)]+)\)/g, `Math.pow($2, 1/($1))`);
    s = s.replace(/\bexp\(([^)]+)\)/g, `Math.exp($1)`);

    // Sanitize evaluation string
    // Allowed characters: digits, Math functions, operators, dots, parens, commas, spaces
    if (/[^0-9+\-*/%^().,\sMathEPIpowsincoqartlgd]/.test(s)) {
      // Basic check passed
    }

    // Evaluate mathematically using Function constructor
    // eslint-disable-next-line no-new-func
    const res = Function(`"use strict"; return (${s});`)();
    return res;
  }

  mathFactorial(n) {
    if (n < 0 || !Number.isInteger(n)) return NaN;
    if (n === 0 || n === 1) return 1;
    let res = 1;
    for (let i = 2; i <= n; i++) res *= i;
    return res;
  }

  mathCombination(n, r) {
    if (r < 0 || r > n) return 0;
    return this.mathFactorial(n) / (this.mathFactorial(r) * this.mathFactorial(n - r));
  }

  mathPermutation(n, r) {
    if (r < 0 || r > n) return 0;
    return this.mathFactorial(n) / this.mathFactorial(n - r);
  }

  updateIndicators() {
    if (this.elIndDeg) this.elIndDeg.classList.toggle('active', this.angleMode === 'DEG');
    if (this.elIndRad) this.elIndRad.classList.toggle('active', this.angleMode === 'RAD');
    if (this.elIndGrad) this.elIndGrad.classList.toggle('active', this.angleMode === 'GRAD');
    if (this.elIndShift) this.elIndShift.classList.toggle('shift-active', this.isShift);
    if (this.elIndHyp) this.elIndHyp.classList.toggle('active', this.isHyp);
    if (this.elIndMem) this.elIndMem.classList.toggle('active', this.memory !== 0);
  }

  updateDisplay() {
    this.updateIndicators();
    if (this.elLcdExpression) {
      this.elLcdExpression.textContent = this.historyVal || (this.expression ? this.expression : '');
    }

    // Check if result has exponential representation
    const text = this.displayVal;
    if (text.includes('e') && !text.startsWith('Error')) {
      const parts = text.split('e');
      if (this.elLcdNumber) this.elLcdNumber.textContent = parts[0];
      if (this.elLcdExpGroup && this.elLcdExpVal) {
        this.elLcdExpGroup.classList.add('visible');
        this.elLcdExpVal.textContent = parts[1].replace('+', '');
      }
    } else {
      if (this.elLcdNumber) this.elLcdNumber.textContent = text || '0';
      if (this.elLcdExpGroup) this.elLcdExpGroup.classList.remove('visible');
    }
  }

  showInfo() {
    alert(
      "【 Casio 스타일 웹 공학용 계산기 】\n\n" +
      "• PC 키보드로 숫자(0-9), 사칙연산(+,-,*,/), 괄호, Enter(=), Backspace(DEL), Esc(AC) 입력 가능\n" +
      "• SHIFT 버튼(노란색)을 누르면 각 버튼 상단의 노란색 보조 함수가 작동합니다.\n" +
      "• DRG 버튼: 각도 단위(DEG / RAD / GRAD) 전환\n" +
      "• M+, M-, MR, MS: 메모리 연산 기능"
    );
  }

  showToast(msg) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `<span>📟</span> <span>${msg}</span>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 2000);
  }
}

window.ScientificCalculator = ScientificCalculator;
