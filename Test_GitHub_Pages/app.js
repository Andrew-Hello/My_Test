let count = 0;
const countEl = document.getElementById('count');
const button = document.getElementById('button');
const reset = document.getElementById('reset');

button.addEventListener('click', () => {
  count += 1;
  countEl.textContent = String(count);
});

reset.addEventListener('click', () => {
  count = 0;
  countEl.textContent = '0';
});
