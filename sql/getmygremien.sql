select 
tgr.grname, grkurz, mgfunk, 
mgadat, 
mgedat, 
RTRIM(pepartei) as Partei, 
rtrim(amname), *
--adname, 
--from tpe
from tmg
join tpe on tmg.mgpenr=tpe.penr
join tgr on tmg.mggrnr=tgr.grnr
join tam on tmg.mgamnr=tam.amnr
--join tad on tpe.peadnr=tad.adnr
where peadnr = 152 
order by tgr.grname,tmg.mgadat
