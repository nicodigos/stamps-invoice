export function isEmailFileItem(item) {
  return /\.(eml|msg)$/i.test(item?.file?.name || "");
}

export function mergeQueuedFileItems(existingItems, incomingItems) {
  let lastIncomingEmailIndex = -1;
  incomingItems.forEach((item, index) => {
    if (isEmailFileItem(item)) lastIncomingEmailIndex = index;
  });

  if (lastIncomingEmailIndex < 0) return [...existingItems, ...incomingItems];

  const retainedExisting = existingItems.filter((item) => !isEmailFileItem(item));
  const retainedIncoming = incomingItems.filter((item, index) => (
    !isEmailFileItem(item) || index === lastIncomingEmailIndex
  ));
  return [...retainedExisting, ...retainedIncoming];
}
