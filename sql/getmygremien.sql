select 
grname, grkurz, mgfunk, 
mgadat, 
mgedat, pepartei
--,
--*
--adname, 
--from tpe
from tmg
join tpe on tmg.mgpenr=tpe.penr
join tgr on tmg.mggrnr=tgr.grnr
-- join tgr on tgr.grnr=tmg.mggrnr
--join tad on tpe.peadnr=tad.adnr
where peadnr = 152 
order by grname,mgadat
--and mgedat = 0
--order by tad.adname