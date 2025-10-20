const clone = (value) => {
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
};

const createStore = (initialState) => {
  let state = clone(initialState);
  const listeners = new Set();

  const getState = () => clone(state);

  const setState = (updater) => {
    const nextState = typeof updater === "function" ? updater(clone(state)) : updater;
    state = { ...state, ...nextState };
    listeners.forEach((listener) => listener(getState()));
  };

  const subscribe = (listener) => {
    listeners.add(listener);
    listener(getState());
    return () => listeners.delete(listener);
  };

  return { getState, setState, subscribe };
};

export { createStore };
