Macro "PuntosArcosCortadosDir"
	userprofiledir = GetEnvironmentVariable("USERPROFILE")
	if userprofiledir <> "" then do
		userprofiledir = userprofiledir + "\\" + "My Documents"
	end
	
	file = ChooseFile({
     {"Geographic File (*.dbd)", "*.dbd"}},
      "Choose a Geographic File", { {"Initial Directory", userprofiledir} } )
    
    RunMacro ("PuntosArcosCortadosFile", file)    

endMacro

Macro "PuntosArcosCortadosFile"  (dbd_file)
	//
	info = GetDBInfo(dbd_file)
	layinfo = GetDBLayers(dbd_file)
	scope = info[1]
	// Create a map using this scope
	map = CreateMap("Para Cortar", {
     {"Scope", scope},
     })
	layer_lineas = AddLayer(map, layinfo[2], dbd_file, layinfo[2], {{"Shared", "True"}})
	RunMacro("G30 new layer default settings", layer_lineas)
    SetLayer(layer_lineas)
	//
	SetMapUnits("Meters")
	lay = GetLayer()
	inf = GetLayerInfo(lay)
	dbd_file = inf[10]
	sinextension = Substring(dbd_file, 1, StringLength(dbd_file) - 4)
    csv_file = sinextension + "_pts.csv"
    fptr2 = OpenFile(csv_file, "w")
    WriteLine(fptr2, "ID,ID_ARCO,DIR,SEC_PTO,X,Y")
    idn = 0
	maxdist = 130
	lay = GetLayer()
	layer_type = GetLayerType(lay)
	if (layer_type <> "Line") then do
		ShowMessage("Debe seleccionar capa de arcos")
		Return()
	end
	SetLayer(lay)
	rec = GetFirstRecord(lay + "|", null)
	while rec <> null do
		lay_id = RH2ID(rec)
        sec = 0
		pts = GetLine(lay_id)
		for i = 1 to pts.length - 1 do
			loc1 = pts[i]
			loc2 = pts[i + 1]
            sec = sec + 1
            idn = idn + 1
            WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",1" + "," + String(sec) + "," + Format(loc1.Lon/1000000, "*.000000") + "," + Format(loc1.Lat/1000000, "*.000000"))
            dist = GetDistance(loc1, loc2)
			if dist > maxdist then do
				difx = loc2.Lon - loc1.Lon
				dify = loc2.Lat - loc1.Lat
				st = Floor(dist / maxdist)
				locax = loc1.Lon
				locay = loc1.Lat
                for j = 1 to st do
					prog = (j * maxdist) / dist
					locnx = loc1.Lon + difx * prog
					locny = loc1.Lat + dify * prog
                    sec  = sec + 1
					idn = idn + 1
                    WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",1" + "," + String(sec) + "," + Format(locnx/1000000, "*.000000") + "," + Format(locny/1000000, "*.000000"))
                    locax = locnx
					locay = locny
				end
			end
			else do
				locax = loc1.Lon
				locay = loc1.Lat
			end
			if loc2.Lon <> locax or loc2.Lat <> locay then do
				idn = idn + 1
                sec = sec  + 1
                WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",1" + "," + String(sec) + "," + Format(loc2.Lon/1000000, "*.000000") + "," + Format(loc2.Lat/1000000, "*.000000"))
            end
		end
		for i = pts.length to 2 step -1 do
			loc1 = pts[i]
			loc2 = pts[i - 1]
            sec = sec + 1
            idn = idn + 1
            WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",-1" + "," + String(sec) + "," + Format(loc1.Lon/1000000, "*.000000") + "," + Format(loc1.Lat/1000000, "*.000000"))
            dist = GetDistance(loc1, loc2)
			if dist > maxdist then do
				difx = loc2.Lon - loc1.Lon
				dify = loc2.Lat - loc1.Lat
				st = Floor(dist / maxdist)
				locax = loc1.Lon
				locay = loc1.Lat
                for j = 1 to st do
					prog = (j * maxdist) / dist
					locnx = loc1.Lon + difx * prog
					locny = loc1.Lat + dify * prog
                    sec  = sec + 1
					idn = idn + 1
                    WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",-1" + "," + String(sec) + "," + Format(locnx/1000000, "*.000000") + "," + Format(locny/1000000, "*.000000"))
                    locax = locnx
					locay = locny
				end
			end
			else do
				locax = loc1.Lon
				locay = loc1.Lat
			end
			if loc2.Lon <> locax or loc2.Lat <> locay then do
				idn = idn + 1
                sec = sec  + 1
                WriteLine(fptr2, String(idn) +  "," + IntToString(lay_id) + ",-1" + "," + String(sec) + "," + Format(loc2.Lon/1000000, "*.000000") + "," + Format(loc2.Lat/1000000, "*.000000"))
            end
		end
		rec = GetNextRecord(lay + "|", rec, null)
	end
	CloseFile(fptr2)
    //
	CloseMap(map)
	//
endMacro
