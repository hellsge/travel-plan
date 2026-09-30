(()=>{
  const today=new Date(),local=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  const labels={ongoing:'进行中',upcoming:'即将出发',past:'已结束'};
  const list=document.getElementById('tripList'),featured=document.getElementById('featuredList'),featuredSection=document.getElementById('featuredSection');
  const cards=[...list.querySelectorAll('.trip-card')];
  const phaseOf=card=>card.dataset.end<local?'past':(card.dataset.start>local?'upcoming':'ongoing');
  const active=[];
  cards.forEach(card=>{
    const phase=phaseOf(card),badge=card.querySelector('.phase');
    badge.textContent=labels[phase];badge.className=`phase phase-${phase}`;
    if(phase!=='past'){card.classList.add('featured');active.push(card)}
  });
  active.sort((a,b)=>a.dataset.start.localeCompare(b.dataset.start)).forEach(card=>featured.append(card));
  featuredSection.hidden=!active.length;
})();
