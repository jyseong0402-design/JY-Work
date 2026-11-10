// 선 아이콘 (상단 도구줄용). currentColor를 따라 색이 바뀝니다.
const P = {
  back: '<path d="M19 12H5"/><path d="M11 5l-7 7 7 7"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
  pen: '<path d="M4 20l1.2-4.8L16 4.4a2 2 0 0 1 2.8 0l.8.8a2 2 0 0 1 0 2.8L8.8 18.8z"/><path d="M14.5 6l3.5 3.5"/>',
  hl: '<path d="M9.5 11.5l6.4-6.4a2 2 0 0 1 2.8 0l.2.2a2 2 0 0 1 0 2.8l-6.4 6.4"/><path d="M9.5 11.5L7 14v3h3l2.5-2.5z"/><path d="M4 21h9"/>',
  eraser: '<path d="M8.5 19.5l-4-4a2 2 0 0 1 0-2.8l8.2-8.2a2 2 0 0 1 2.8 0l4 4a2 2 0 0 1 0 2.8L12 19.5z"/><path d="M9 9.5l6 6"/><path d="M12 19.5h8"/>',
  text: '<path d="M5 7V5h14v2"/><path d="M12 5v14"/><path d="M9 19h6"/>',
  ai: '<path d="M12 3.5l1.7 4.6 4.6 1.7-4.6 1.7L12 16.1l-1.7-4.6-4.6-1.7 4.6-1.7z"/><path d="M18.5 15l.7 1.8 1.8.7-1.8.7-.7 1.8-.7-1.8-1.8-.7 1.8-.7z"/>',
  info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5"/><path d="M12 7.6v.4"/>',
  full: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
  more: '<circle cx="12" cy="5.5" r="1.4" fill="currentColor"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/><circle cx="12" cy="18.5" r="1.4" fill="currentColor"/>',
};
export const icon = (name) =>
  `<svg class="ico" viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[name] || ''}</svg>`;
