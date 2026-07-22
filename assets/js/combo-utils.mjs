export function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

export function filterOptionsContaining(options, query) {
  const normalizedQuery = normalizeSearchText(query);
  if (!normalizedQuery) return [...options];
  return options.filter((option) => normalizeSearchText(option.label).includes(normalizedQuery));
}
