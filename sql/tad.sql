select tad.adname, tad.advname, tad.adannr,
	 tad.*, tpe.*, tsb.*, tat.*, tan.* from tad
				inner join tan on tan.annr = tad.adannr
                inner join tpe on tad.adnr = tpe.peadnr
                inner join tsb on tsb.sbnr = tpe.pesbnr
                inner join tat on tat.atnr = tsb.sbatnr
				
                where peedat = 0 and adnr > 0 and adname like 'Waller'
				order by tad.adname
				