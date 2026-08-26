/**
 * Main Application Controller: Responsive Tabs, Focus Management, and Touch Swipe
 */

document.addEventListener('DOMContentLoaded', () => {
  // Global Active Panel State ('calc' or 'sheet')
  window.activePanel = 'calc';

  const calcPanel = document.getElementById('calcPanel');
  const sheetPanel = document.getElementById('sheetPanel');
  const calcBadge = document.getElementById('calcBadge');
  const sheetBadge = document.getElementById('sheetBadge');

  window.setActivePanel = function(panel) {
    window.activePanel = panel;
    if (panel === 'calc') {
      if (calcPanel) calcPanel.classList.add('is-active');
      if (sheetPanel) sheetPanel.classList.remove('is-active');
      if (calcBadge) calcBadge.textContent = '🟢 활성화됨';
      if (sheetBadge) sheetBadge.textContent = '클릭하여 활성화';
    } else {
      if (sheetPanel) sheetPanel.classList.add('is-active');
      if (calcPanel) calcPanel.classList.remove('is-active');
      if (sheetBadge) sheetBadge.textContent = '🟢 활성화됨';
      if (calcBadge) calcBadge.textContent = '클릭하여 활성화';
    }
  };

  // Initialize Components
  window.calculatorInstance = new ScientificCalculator();
  window.spreadsheetInstance = new Spreadsheet(20, 20);

  // Set initial active state
  window.setActivePanel('calc');

  // Mobile Tab Elements
  const tabCalc = document.getElementById('tabCalc');
  const tabSheet = document.getElementById('tabSheet');
  const slidingBar = document.getElementById('tabSlidingBar');
  const mainContainer = document.getElementById('mainContainer');

  let currentTab = 'calc';

  function switchTab(target) {
    currentTab = target;
    window.setActivePanel(target);

    if (target === 'calc') {
      tabCalc.classList.add('active');
      tabSheet.classList.remove('active');
      slidingBar.style.transform = 'translateX(0%)';
      mainContainer.classList.remove('tab-sheet');
      mainContainer.classList.add('tab-calc');
    } else {
      tabSheet.classList.add('active');
      tabCalc.classList.remove('active');
      slidingBar.style.transform = 'translateX(100%)';
      mainContainer.classList.remove('tab-calc');
      mainContainer.classList.add('tab-sheet');
    }
  }

  if (tabCalc && tabSheet) {
    tabCalc.addEventListener('click', () => switchTab('calc'));
    tabSheet.addEventListener('click', () => switchTab('sheet'));
  }

  // Mobile Touch Swipe Navigation
  let touchStartX = 0;
  let touchStartY = 0;
  let touchEndX = 0;
  let touchEndY = 0;

  mainContainer.addEventListener('touchstart', (e) => {
    touchStartX = e.changedTouches[0].screenX;
    touchStartY = e.changedTouches[0].screenY;
  }, { passive: true });

  mainContainer.addEventListener('touchend', (e) => {
    touchEndX = e.changedTouches[0].screenX;
    touchEndY = e.changedTouches[0].screenY;
    handleSwipe();
  }, { passive: true });

  function handleSwipe() {
    const diffX = touchEndX - touchStartX;
    const diffY = touchEndY - touchStartY;
    
    if (Math.abs(diffX) > Math.abs(diffY) * 1.5 && Math.abs(diffX) > 60) {
      if (diffX < 0 && currentTab === 'calc') {
        switchTab('sheet');
      } else if (diffX > 0 && currentTab === 'sheet') {
        switchTab('calc');
      }
    }
  }
});
