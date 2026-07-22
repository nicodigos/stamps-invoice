export function singleFlight(work) {
  let inProgress = false;

  return async (...args) => {
    if (inProgress) return false;
    inProgress = true;
    try {
      await work(...args);
      return true;
    } finally {
      inProgress = false;
    }
  };
}
