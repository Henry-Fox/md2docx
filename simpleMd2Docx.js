createListFromMarked(markedArray) {
  // 在展开操作前添加类型检查
  const safeArray = Array.isArray(markedArray) ? markedArray : [];

  // 修改前：可能直接使用 ...markedArray
  // 修改后：使用保护后的 safeArray
  return [
    /* ... other content ... */,
    ...safeArray   // 将safeArray的所有元素展开到新数组中
  ];
}
