(()=>{
  const copyText=async button=>{
    try{await navigator.clipboard.writeText(button.dataset.copy)}catch(_){
      const area=document.createElement('textarea');area.value=button.dataset.copy;document.body.append(area);area.select();document.execCommand('copy');area.remove();
    }
    const old=button.textContent;button.textContent='已复制';setTimeout(()=>button.textContent=old,1200);
  };
  const mapDialog=document.getElementById('mapDialog'),mapPlace=document.getElementById('mapPlace');let selectedPlace='';
  const mapUrl=(provider,place)=>{
    const encoded=encodeURIComponent(place);
    if(provider==='apple')return `https://maps.apple.com/?q=${encoded}`;
    if(provider==='baidu')return `https://api.map.baidu.com/geocoder?address=${encoded}&output=html&src=travel-plan`;
    return `https://uri.amap.com/search?keyword=${encoded}`;
  };
  document.addEventListener('click',event=>{
    const copy=event.target.closest('.copy,.copy-item');if(copy){copyText(copy);return}
    const opener=event.target.closest('.map-open');if(!opener)return;
    selectedPlace=opener.dataset.location;mapPlace.textContent=selectedPlace;
    const preferred=localStorage.getItem('preferredMap')||'amap';
    mapDialog.querySelectorAll('[data-map]').forEach(button=>button.classList.toggle('primary',button.dataset.map===preferred));
    if(mapDialog.showModal)mapDialog.showModal();else window.open(mapUrl(preferred,selectedPlace),'_blank','noopener');
  });
  mapDialog.querySelectorAll('[data-map]').forEach(button=>button.addEventListener('click',()=>{
    localStorage.setItem('preferredMap',button.dataset.map);window.open(mapUrl(button.dataset.map,selectedPlace),'_blank','noopener');mapDialog.close();
  }));
  document.getElementById('mapClose').addEventListener('click',()=>mapDialog.close());
  mapDialog.addEventListener('click',event=>{if(event.target===mapDialog)mapDialog.close()});
  const days=[...document.querySelectorAll('.day')],links=[...document.querySelectorAll('[data-day-link]')],toolbar=document.querySelector('.toolbar'),today=new Date(),local=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  let current=days.find(day=>day.dataset.date===local)||days.find(day=>day.dataset.date>local)||days[days.length-1],scrollTarget=null,scrollTimer=0,scrollFrame=0;
  const setActive=date=>links.forEach(link=>link.classList.toggle('active',link.dataset.dayLink===date));
  if(current){current.classList.add('today');setActive(current.dataset.date)}
  const markCurrentItem=()=>{
    document.querySelectorAll('.item.current,.item.next').forEach(item=>item.classList.remove('current','next'));
    const now=new Date(),items=[...document.querySelectorAll('.item[data-start]')].filter(item=>item.dataset.start&&!Number.isNaN(new Date(item.dataset.start).getTime()));
    const active=items.find(item=>{const start=new Date(item.dataset.start),end=item.dataset.end?new Date(item.dataset.end):new Date(start.getTime()+60*60*1000);return start<=now&&now<end});
    const next=items.find(item=>new Date(item.dataset.start)>now);
    if(active)active.classList.add('current');else if(next)next.classList.add('next');
  };
  markCurrentItem();setInterval(markCurrentItem,60000);
  const stickyOffset=()=>toolbar.offsetHeight+(parseFloat(getComputedStyle(toolbar).top)||0)+8;
  if(!location.hash){
    const anchor=document.querySelector('.item.current,.item.next')||(current&&current.querySelector('.day-header'));
    if(anchor){const top=window.scrollY+anchor.getBoundingClientRect().top-stickyOffset();window.scrollTo({top:Math.max(0,top),behavior:'smooth'})}
  }
  const syncActiveDay=()=>{
    scrollFrame=0;if(scrollTarget)return;
    const visibleDays=days.filter(day=>!day.hidden);if(!visibleDays.length)return;
    let active=visibleDays[0],line=stickyOffset();
    if(window.scrollY+window.innerHeight>=document.documentElement.scrollHeight-2)active=visibleDays[visibleDays.length-1];
    else for(const day of visibleDays){if(day.getBoundingClientRect().top<=line)active=day;else break}
    setActive(active.dataset.date);
  };
  const finishNavigation=()=>{if(!scrollTarget)return;const date=scrollTarget;scrollTarget=null;setActive(date);syncActiveDay()};
  links.forEach(link=>link.addEventListener('click',event=>{
    event.preventDefault();const target=document.getElementById(`day-${link.dataset.dayLink}`);if(!target||target.hidden)return;
    scrollTarget=link.dataset.dayLink;setActive(scrollTarget);clearTimeout(scrollTimer);
    history.replaceState(null,'',`#day-${scrollTarget}`);
    const top=window.scrollY+target.getBoundingClientRect().top-stickyOffset();
    window.scrollTo({top:Math.max(0,top),behavior:'smooth'});
    scrollTimer=setTimeout(finishNavigation,1200);
  }));
  window.addEventListener('scroll',()=>{
    if(scrollTarget){clearTimeout(scrollTimer);scrollTimer=setTimeout(finishNavigation,160);return}
    if(!scrollFrame)scrollFrame=requestAnimationFrame(syncActiveDay);
  },{passive:true});
})();
