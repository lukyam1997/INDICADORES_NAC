const createElement = (tag, options = {}) => {
  const el = document.createElement(tag);
  if (options.className) {
    el.className = options.className;
  }
  if (options.text) {
    el.textContent = options.text;
  }
  if (options.html) {
    el.innerHTML = options.html;
  }
  if (options.attrs) {
    Object.entries(options.attrs).forEach(([key, value]) => {
      el.setAttribute(key, value);
    });
  }
  if (options.children) {
    options.children.forEach((child) => el.append(child));
  }
  return el;
};

const clearElement = (el) => {
  while (el.firstChild) {
    el.removeChild(el.firstChild);
  }
};

export { createElement, clearElement };
