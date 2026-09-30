(()=>{
  if(location.search.includes('_t='))history.replaceState(null,'',location.pathname+location.hash);
  const themeToggle=document.getElementById('themeToggle');
  const applyTheme=dark=>{
    document.documentElement.dataset.theme=dark?'dark':'light';
    themeToggle.textContent=dark?'☀️':'🌙';
    localStorage.setItem('theme',dark?'dark':'light');
  };
  const stored=localStorage.getItem('theme');
  applyTheme(stored?stored==='dark':matchMedia('(prefers-color-scheme: dark)').matches);
  themeToggle.addEventListener('click',()=>applyTheme(document.documentElement.dataset.theme!=='dark'));
  const offlineBadge=document.getElementById('offlineBadge'),updateNetwork=()=>offlineBadge.hidden=navigator.onLine;
  window.addEventListener('online',updateNetwork);window.addEventListener('offline',updateNetwork);updateNetwork();
  if('serviceWorker' in navigator && location.protocol.startsWith('http')){
    window.addEventListener('load',()=>{
      navigator.serviceWorker.register(document.documentElement.dataset.root+'service-worker.js',{updateViaCache:'none'}).catch(error=>console.warn('离线缓存注册失败',error));
    });
  }
})();
