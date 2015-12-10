
Macro "ProcesarBatchCsvGps"

	userprofiledir = GetEnvironmentVariable("USERPROFILE")
	if userprofiledir <> "" then do
		userprofiledir = userprofiledir + "\\" + "My Documents"
	end
	
	dir = ChooseDirectory("Elija la carpeta raíz con los archivos .geo", { {"Initial Directory", userprofiledir} })
	
	if Substring(dir, StringLength(dir), 1) = "\\" then do
		dir = Substring(dir, 1, StringLength(dir) - 1)
	end
	
	di = GetDirectoryInfo(dir + "\\" + "*.*", "Directory")

	for i = 1 to di.length do

		sub_dir = dir + "\\" + di[i][1]
		
		sdi = GetDirectoryInfo(sub_dir + "\\" + "*.*", "Directory")
	
		for j = 1 to sdi.length do

			ssub_dir = sub_dir + "\\" + sdi[j][1]
			
			ssdi = GetDirectoryInfo(ssub_dir + "\\" + "*.geo", "File")

			for k = 1 to ssdi.length do

				file = ssub_dir + "\\" + ssdi[k][1]
			
				RunMacro ("ImportCsvGps", file)
		
			end
		
		end
		
    end
	
endMacro

Macro "ImportCsvGpsItfc"
	file = ChooseFile({
     {"Text geographic (*.geo)", "*.geo"}},
      "Choose a Geographic Text File", )
	RunMacro ("ImportCsvGps", file)
endMacro

Macro "ImportCsvGps" (geo_file)

	geo_sinext = Substring(geo_file, 1, StringLength(geo_file) - 4)
	csv_file = geo_sinext + ".csv"
	csx_file = geo_sinext + ".csx"
	dbd_file = geo_sinext + ".dbd"
	dbd_tmp_file = geo_sinext + "_tmp"
	dbd_tmp_file = dbd_tmp_file + ".dbd"
	solonombre = ""
	RunMacro ("SoloNombre", dbd_file, &solonombre)

	ImportCSV(geo_file, dbd_tmp_file, "Line", {
	, {"Direction", 2}
	, {"Geography", 3}
	, {"ID", 1}
	, {"Label", solonombre}
	, {"Layer Name", solonombre}
	, {"Node Layer Name", "Ep " + solonombre}
	})

	// Get the scope of a geographic file

/*
	layer_lineas = AddLayer(null, solonombre, dbd_tmp_file, solonombre, {{"Shared", "True"}})
	RunMacro("G30 new layer default settings", layer_lineas)
	 
	layer_nodos = AddLayer(null, "Ep " + solonombre, dbd_tmp_file, "Ep " + solonombre, {{"Shared", "True"}})
	RunMacro("G30 new layer default settings", layer_nodos)

	SetLayerVisibility(layer_nodos, "False")
*/
	 
	JoinTableToLayer(dbd_tmp_file, solonombre, "CSV", csv_file, , "ID", {{"Hide Link", "True"}})

	info = GetDBInfo(dbd_tmp_file)
	scope = info[1]
	// Create a map using this scope
	map = CreateMap(solonombre, {
     {"Scope", scope},
     })

	layer_lineas = AddLayer(null, solonombre, dbd_tmp_file, solonombre, {{"Shared", "True"}})
	RunMacro("G30 new layer default settings", layer_lineas)

	fields_array = GetFields(layer_lineas, "All")
	field_specs = fields_array[2]
	
	fields_array = GetMappableFields(layer_lineas, "All")
	field_specs = fields_array[2]
	
	ExportGeography(layer_lineas, dbd_file, { {"Internal Data", "True"}, {"Field Spec", field_specs} } )
	
	DropLayer(null, layer_lineas)

	DeleteDatabase(dbd_tmp_file)
	DeleteFile(csx_file)
	
	CloseMap(map)
	
	//view = OpenTable(solonombre + " CSV", "CSV", {csv_file, null}, {{"Shared", "True"}})
	//closeview(view)
	
endMacro

Macro "SoloNombre" (nombrecompleto, solonombre)
	start = 1
	pos = 0
	aux = nombrecompleto
	pos = PositionFrom(start, aux, "\\")
	while pos > 0 do
		start = pos + 1
		aux = nombrecompleto
		pos = PositionFrom(start, aux, "\\")
    end
	aux = Substring(nombrecompleto, start, StringLength(nombrecompleto) - start + 1)
	if Position(aux, ".") > 0 then do
		aux = Substring(aux, 1, Position(aux, ".") - 1)
	end
	solonombre = aux
endMacro
