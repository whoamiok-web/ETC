// 우리 가족 달력: 알림 창을 누르면 앱을 열어 줌 (화면 캐시는 하지 않음)
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    for (const c of list) if (c.url.includes('family_calendar')) return c.focus();
    return self.clients.openWindow('family_calendar.html');
  }));
});
