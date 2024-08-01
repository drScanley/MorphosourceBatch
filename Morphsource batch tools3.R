#'@export
#'Extract CT metadata and specimen information from files and create a Morphosource Batch manifest
#'@param DIRECTORY path to the directory with your files.
#'@param Institution Darwin core code of the institution where your specimens are from.
#'@param CollectionCode Darwin core Code of the collection where your specimens are from.
#'@param Creator Individual or individuals that produced the original data
#'@param Scanning.system Brand of CT scanner used to produce metadata files if \code{"Bruker"} Bruker log files. if \code{"Nikon"} Nikon xtekct files, if \code{"Phoenix"}Phoenix .pca files
#'@param Delimiter characters used to split Darwin core Triplets. if \code{D1} = spaces,hyphens and underscores, if \code{D2} = Spaces and underscores if \code{D3} = Underscores and hyphens, if \code{D4} = spaces and hyphens)
#'@param Name.system How to parse file names if \code{"3Part"} = Inst:Col:SpecNo (e.g.UF-Herp-12345) if \code{"2Part"} = INST:SpecNo (e.g. UF-12345)
#'@param AddAccessionFirst Included if the specimen number requires a leading qualifier (e.g. AddAccessionFirst="R" = MCZ_herps_R12345)
#'@param AddAccessionLast Included if the specimen number requires a trailing qualifier (e.g. AddAccessionLast="-frog" = MCZ_herps_12345-frog)
#'@param media_pub_status "Private":"Restricted download":"Open download"
#'@param media_type pre-set type of data
#'@param media_raw_or_derived whether the media / zip files contains raw— i.e. radiographs— or derived —ie tomograms, mesh files, videos etc
#'@return a formatted xlsx file  for Morphosource batch upload
#'@export
#'@examples 
#'MSMetadataBatch<-function(DIRECTORY="UF",Institution="UF", CollectionCode="herp",Creator="Edward Stanley",Scanning.system="Phoenix",Delimiter="D1",Name.system="3Part",AddAccession="",media_pub_status="private",media_type="CTImageSeries",media_raw_or_derived="derived")



  



MSMetadataBatch<-function(DIRECTORY,Institution="UF", CollectionCode="herp",Creator="Edward Stanley",Scanning.system="Phoenix",Delimiter="D1",Name.system="3Part",AddAccessionFirst="",AddAccessionLast="",media_pub_status="private",media_type="CTImageSeries",media_raw_or_derived="derived", exposure="Unreported")
{
require(ridigbio)
require(stringr)
require(readr)
require(xlsx)
#remember current working directory
mypath <- getwd()

#set standard values
media_type<-"CTImageSeries"
media_raw_or_derived<-"derived"  
media.series_type <-"reconstructed image stack"
media.unit<-"mm"
media.part<-"entire specimen"
media.preview_file<-""
media.parent_file<-""
media.parent_ms_id<-""
biological_specimen.ms_id<-""
biological_specimen.occurrence_id<-""
media.short_description<-"microCT volume and derivatives"
media.side<-"NotApplicable"
media.description<-""
media.orientation<-""
media.identifier<-""
media.keyword<-""
media.date_created<-""
media.related_url<-""
media.map_type<-""
biological_specimen.vouchered<-"yes"
biological_specimen.identifier<-""
biological_specimen.related_url<-""
biological_specimen.description<-""
biological_specimen.numeric_time<-""
biological_specimen.original_location<-""
biological_specimen.periodic_time<-""
biological_specimen.is_type_specimen<-""
biological_specimen.sex<-""
taxonomy.taxonomy_subspecies<-""
Avg<-"unreported"
TimingVal<-"unreported"
imaging_event.description	<-""
imaging_event.software	<-""
imaging_event.date_created <-""
imaging_event.ct.flux_normalization	<-"no"
imaging_event.ct.geometric_calibration	<-""
imaging_event.ct.shading_correction<-"no"
imaging_event.ct.surrounding_material<-"air"
imaging_event.ct.xray_tube_type<-""
imaging_event.ct.target_type<-"reflection"
imaging_event.ct.detector_type<-" Scintillator (Phosphor used)"
imaging_event.ct.detector_configuration <-"Area (single or tiled detector)"
imaging_event.ct.target_material<-"tungston"
imaging_event.ct.rotation_number<-""
imaging_event.ct.phase_contrast<-"no"
imaging_event.ct.optical_magnification<-"no"
imaging_event.ct.acquisition_type	<-""
imaging_event.photogrammetry.focal_length_type	<-""
imaging_event.photogrammetry.background_removal	<-""
imaging_event.photography.lens_make	<-""
imaging_event.photography.lens_model	<-""
imaging_event.photography.light_source	<-""
processing_event.date_created	<-""
processing_event.software	<-"DatosX|R"
processing_event.description <-""
processing_event.creator<-Creator
#D1 = Spaces, hyphens and underscores
if(Delimiter=="D1"){
  Delimiter1<-"[\\ |,_,-]+"
}
#D2 = Spaces and underscores
else if(Delimiter=="D2"){
    Delimiter1<-"[\\ |,_]+"
  }
#D3 = hyphens underscores
else if(Delimiter=="D3"){
  Delimiter1<-"[,_,-]+"
}
#D4 = spaces and hyphens
else  if(Delimiter=="D4"){
    Delimiter1<-"[\\ |,-]+"
  }
else{
  Delimiter1<-Delimiter
}



###
######
### If scanning on a Phoenix system with PCA metadata files
######
###

if(Scanning.system=="Phoenix"){
#get the names and number of files in DIRECTORY for loop
list.files(path = DIRECTORY,pattern = ".pca")-> pcalist
length(pcalist)->metadatafileNo
#make a blank data frame to be filled in by loop
Matrix1<-data.frame()
for (i in 1:metadatafileNo){
  
#set up text reporting for loop
  Sys.sleep(0.1)
  round(i/metadatafileNo*100,digits=2)->P
  
  #Get file in DIRECTORY
  pcalist[i]->DWCTrip
  #report which file is being processed
  message(paste(P,"%","parsing",DWCTrip, sep=" "))
  #Change .pca to .zip for media file
  DWCTripZip<-gsub("pca","zip",as.character(DWCTrip))
  #remove .pca and split remaining file name by -, _ or space
  DWCTrip<-gsub(".pca","",as.character(DWCTrip))
  as.list(unlist(strsplit(DWCTrip, split=Delimiter1)))->DWCTList
  if(Name.system=="3Part"){
    #return media part and accession number
    AccessionNo<-unlist(DWCTList[3])
    AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
    if(nchar(AddAccessionFirst) >=1){
      AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
    if(nchar(AddAccessionLast) >=1){
      AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
    #return 4th component as media part (if it exists)
    if(length(unlist(DWCTList))>3)
    {
      media.part<-unlist(DWCTList[4])
    }else{
      media.part<-"entire specimen"
    }
  }
  if(Name.system=="2Part"){
    AccessionNo<-unlist(DWCTList[2])
    AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
    if(nchar(AddAccessionFirst) >=1){
      AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
    if(nchar(AddAccessionLast) >=1){
      AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
    #return 4th component as media part (if it exists)
    if(length(unlist(DWCTList))>2)
    {
      media.part<-unlist(DWCTList[3])
      
    }else{
      
      media.part<-"entire specimen"
    }
  }
  #ping idigbio for Institution, collection and accessionno
  Search1 <- list("scientificname"=list("type"="exists"), 
                  "catalognumber"=AccessionNo,"institutioncode"=Institution,"collectioncode"=CollectionCode)
  ridigbio::idig_search_records(rq = Search1,limit = 11)->results1
 #Pull out relevant info from iDigBio if found
  if(nrow(results1)>=1){
  UUID<-results1$uuid
  occurrenceID<-results1$occurrenceid
  ScientificName<-results1$scientificname
  Genus<-results1$genus
  Genus<-str_to_title(Genus)
  SpeciesName<-gsub(results1$genus,"",as.character(results1$scientificname))
  DateCollected<-results1$`data.dwc:eventDate`
  Collector<-results1$collector
  Lattitude<-results1$geopoint.lat
  Longitude<-results1$geopoint.lon
  
  #report that the search didnt find anything on idigbio
  }else{
    UUID<-""
    occurrenceID<-""
    ScientificName<-""
    Genus<-""
    SpeciesName<-"Not on IdigBio"
    DateCollected<-""
    Collector<-"Not on IdigBio"
    Lattitude<-""
    Longitude<-""
  }
  #open pca file and read data into relevant objects
  paste(DIRECTORY,pcalist[i],sep = "/")->PCApath
  conn<-file(PCApath,open="r")
  linn<-readLines(conn)
  close(conn)
  linn[grep('^VoxelSizeX.*', linn)]->Voxelsize
  Voxelsize<-gsub("VoxelSizeX=","",as.character(Voxelsize))
  Voxelsize<-as.numeric(Voxelsize)
  linn[grep('^NumberImages*', linn)]->ImageNo
  ImageNo<-gsub("NumberImages=","",as.character(ImageNo))
  ImageNo<-as.numeric(ImageNo[1])
  linn[grep('^Voltage=*', linn)]->Voltage
  Voltage<-gsub("Voltage=","",as.character(Voltage))
  Voltage<-as.numeric(Voltage)
  
  linn[grep('^Current=*', linn)]->Current
  Current<-gsub("Current=","",as.character(Current))
  Current<-as.numeric(Current)
  
  linn[grep('^Filter=*', linn)]->Filter
  Filter<-gsub("Filter=","",as.character(Filter))
  if(Filter=="Unknown"){
    Filter<-"None"
  }
  if(Filter=="Please define a filter."){
    Filter<-"None"
  }
  linn[grep('^Avg=*', linn)]->Avg
  Avg<-gsub("Avg=","",as.character(Avg))
  Avg<-as.numeric(Avg[2])
  
  linn[grep('^TimingVal=*', linn)]->TimingVal
  TimingVal<-gsub("TimingVal=","",as.character(TimingVal))
  TimingVal<-as.numeric(TimingVal)
  
  linn[grep('^FDD=*', linn)]->FDD
  FDD<-gsub("FDD=","",as.character(FDD))
  FDD<-as.numeric(FDD)

  linn[grep('^FOD=*', linn)]->FOD
  FOD<-gsub("FOD=","",as.character(FOD))
  FOD<-as.numeric(FOD)
  
  linn[grep('^PixelsizeX=*', linn)]->PixelsizeX
  PixelsizeX<-gsub("PixelsizeX=","",as.character(PixelsizeX))
  PixelsizeX<-as.numeric(PixelsizeX)
  linn[grep('^NrPixelsX=*', linn)]->NrPixelsX
  NrPixelsX<-gsub("NrPixelsX=","",as.character(NrPixelsX))
  NrPixelsX<-as.numeric(NrPixelsX)
  linn[grep('^PixelsizeY=*', linn)]->PixelsizeY
  PixelsizeY<-gsub("PixelsizeY=","",as.character(PixelsizeY))
  PixelsizeY<-as.numeric(PixelsizeY)
  linn[grep('^NrPixelsY=*', linn)]->NrPixelsY
  NrPixelsY<-gsub("NrPixelsY=","",as.character(NrPixelsY))
  NrPixelsY<-as.numeric(NrPixelsY)
(Voltage*Current/1000000)->Wattage

#bind all data into a row
 Dataline <-cbind(DWCTripZip, media.preview_file,  media_type, media_raw_or_derived, media.parent_file, media.parent_ms_id, biological_specimen.ms_id, UUID, occurrenceID,  Institution, CollectionCode, AccessionNo, media.part, media.short_description, media.side, media.description, Creator, media.orientation, media.identifier, media.keyword, media.date_created, media.related_url, Voxelsize, Voxelsize, Voxelsize, Voxelsize, media.series_type, media.unit, media.map_type, biological_specimen.identifier, biological_specimen.related_url, DateCollected, Collector, biological_specimen.description, Lattitude, Longitude, biological_specimen.numeric_time, biological_specimen.original_location, biological_specimen.periodic_time, biological_specimen.is_type_specimen, biological_specimen.sex, biological_specimen.vouchered, Genus, SpeciesName, taxonomy.taxonomy_subspecies, imaging_event.description, Creator, imaging_event.software, imaging_event.date_created, TimingVal, imaging_event.ct.flux_normalization, imaging_event.ct.geometric_calibration, imaging_event.ct.shading_correction, Filter, Avg, ImageNo, Voltage, Wattage, Current, imaging_event.ct.surrounding_material, imaging_event.ct.xray_tube_type, imaging_event.ct.target_type, imaging_event.ct.detector_type, NrPixelsX, PixelsizeX, NrPixelsY, PixelsizeY, imaging_event.ct.detector_configuration, FOD, FDD, imaging_event.ct.target_material, imaging_event.ct.rotation_number, imaging_event.ct.phase_contrast, imaging_event.ct.optical_magnification, imaging_event.ct.acquisition_type, imaging_event.photogrammetry.focal_length_type, imaging_event.photogrammetry.background_removal, imaging_event.photography.lens_make, imaging_event.photography.lens_model, imaging_event.photography.light_source, processing_event.creator, processing_event.date_created, processing_event.software, processing_event.description)
                  
              
#Append row to dataframe Matrix1
  Matrix1<-rbind(Matrix1,Dataline)
}
}




###
######
### If scanning on a Nikon system with xtekct metadata files
######
###

if(Scanning.system=="Nikon"){
  
  #get the names and number of files in DIRECTORY for loop
  list.files(path = DIRECTORY,pattern = ".xtekct")-> Xtekctlist
  length(Xtekctlist)->metadatafileNo
  #make a blank data frame to be filled in by loop
  Matrix1<-data.frame()
  for (i in 1:metadatafileNo){
    
    #set up text reporting for loop
    Sys.sleep(0.1)
    round(i/metadatafileNo*100,digits=2)->P
    
    #Get file in DIRECTORY
    Xtekctlist[i]->DWCTrip
    #report which file is being processed
    message(paste(P,"%","parsing",DWCTrip, sep=" "))
    #Change .xtekct to .zip for media file
    DWCTripZip<-gsub(".xtekct",".zip",as.character(DWCTrip))
    #remove .xtekct and split remaining file name by -, _ or space
    DWCTrip<-gsub(".xtekct","",as.character(DWCTrip))
    as.list(unlist(strsplit(DWCTrip, split=Delimiter1)))->DWCTList
    if(Name.system=="3Part"){
      #return media part and accession number
      AccessionNo<-unlist(DWCTList[3])
      AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
      if(nchar(AddAccessionFirst) >=1){
        AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
      if(nchar(AddAccessionLast) >=1){
        AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
      #return 4th component as media part (if it exists)
      if(length(unlist(DWCTList))>3)
      {
        media.part<-unlist(DWCTList[4])
      }else{
        media.part<-"entire specimen"
      }
    }
    if(Name.system=="2Part"){
      AccessionNo<-unlist(DWCTList[2])
      AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
      if(nchar(AddAccessionFirst) >=1){
        AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
      if(nchar(AddAccessionLast) >=1){
        AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
      #return 4th component as media part (if it exists)
      if(length(unlist(DWCTList))>2)
      {
        media.part<-unlist(DWCTList[3])
        
      }else{
        
        media.part<-"entire specimen"
      }
    }
    #ping idigbio for Institution, collection and accessionno
    Search1 <- list("scientificname"=list("type"="exists"), 
                    "catalognumber"=AccessionNo,"institutioncode"=Institution,"collectioncode"=CollectionCode)
    ridigbio::idig_search_records(rq = Search1,limit = 11)->results1
    #Pull out relevant info from iDigBio if found
    if(nrow(results1)>=1){
      UUID<-results1$uuid
      occurrenceID<-results1$occurrenceid
      ScientificName<-results1$scientificname
      Genus<-results1$genus
      Genus<-str_to_title(Genus)
      SpeciesName<-gsub(results1$genus,"",as.character(results1$scientificname))
      DateCollected<-results1$`data.dwc:eventDate`
      Collector<-results1$collector
      Lattitude<-results1$geopoint.lat
      Longitude<-results1$geopoint.lon
      
      #report that the search didnt find anything on idigbio
    }else{
      UUID<-""
      occurrenceID<-"Not on IdigBio"
      ScientificName<-"Not on IdigBio"
      Genus<-"Not on IdigBio"
      SpeciesName<-"Not on IdigBio"
      DateCollected<-"Not on IdigBio"
      Collector<-"Not on IdigBio"
      Lattitude<-"Not on IdigBio"
      Longitude<-"Not on IdigBio"
    }
    #open xtekct file and read data into relevant objects
    paste(DIRECTORY,Xtekctlist[i],sep = "/")->Xtekctpath
    conn<-file(Xtekctpath,open="r")
    linn<-readLines(conn)
    close(conn)
    linn[grep('^VoxelSizeX=*', linn)]->Voxelsize
    Voxelsize<-readr::parse_number(Voxelsize)
    linn[grep('^Projections*', linn)]->ImageNo
    ImageNo<-readr::parse_number(ImageNo)
    linn[grep('^XraykV=*', linn)]->Voltage
    Voltage<-readr::parse_number(Voltage)
    
    linn[grep('^XrayuA=*', linn)]->Current
    Current<-readr::parse_number(Current)
    
    linn[grep('^Filter_ThicknessMM=*', linn)]->Filter
    Filter<-gsub("Filter_ThicknessMM=","",as.character(Filter))
    if(Filter=="Unknown"){
      Filter<-"None"
    }
    if(Filter=="Please define a filter."){
      Filter<-"None"
    }
    
    linn[grep('^TimingVal=*', linn)]->TimingVal
    TimingVal<-readr::parse_number(TimingVal)
    TimingVal<-exposure
    
    linn[grep('^SrcToDetector=*', linn)]->FDD
    FDD<-readr::parse_number(FDD)
    
    linn[grep('^SrcToObject=*', linn)]->FOD
    FOD<-readr::parse_number(FOD)
    
    linn[grep('^DetectorPixelSizeX=*', linn)]->PixelsizeX
    PixelsizeX<-readr::parse_number(PixelsizeX)
    linn[grep('^DetectorPixelsX=*', linn)]->NrPixelsX
    NrPixelsX<-readr::parse_number(NrPixelsX)
    linn[grep('^DetectorPixelSizeY=*', linn)]->PixelsizeY
    PixelsizeY<-readr::parse_number(PixelsizeY)
    linn[grep('^DetectorPixelsY=*', linn)]->NrPixelsY
    NrPixelsY<-readr::parse_number(NrPixelsY)
    (Voltage*Current/1000000)->Wattage
    
    #bind all data into a row
    Dataline <-cbind(DWCTripZip, media.preview_file,media_type, media_raw_or_derived, media.parent_file, media.parent_ms_id, biological_specimen.ms_id, occurrenceID, biological_specimen.occurrence_id, Institution, CollectionCode, AccessionNo, media.part, media.short_description, media.side, media.description, Creator, media.orientation, media.identifier, media.keyword, media.date_created, media.related_url, Voxelsize, Voxelsize, Voxelsize, Voxelsize, media.series_type, media.unit, media.map_type, biological_specimen.identifier, biological_specimen.related_url, DateCollected, Collector, biological_specimen.description, Lattitude, Longitude, biological_specimen.numeric_time, biological_specimen.original_location, biological_specimen.periodic_time, biological_specimen.is_type_specimen, biological_specimen.sex, biological_specimen.vouchered, Genus, SpeciesName, taxonomy.taxonomy_subspecies, imaging_event.description, Creator, imaging_event.software, imaging_event.date_created, TimingVal, imaging_event.ct.flux_normalization, imaging_event.ct.geometric_calibration, imaging_event.ct.shading_correction, Filter, Avg, ImageNo, Voltage, Wattage, Current, imaging_event.ct.surrounding_material, imaging_event.ct.xray_tube_type, imaging_event.ct.target_type, imaging_event.ct.detector_type, NrPixelsX, PixelsizeX, NrPixelsY, PixelsizeY, imaging_event.ct.detector_configuration, FOD, FDD, imaging_event.ct.target_material, imaging_event.ct.rotation_number, imaging_event.ct.phase_contrast, imaging_event.ct.optical_magnification, imaging_event.ct.acquisition_type, imaging_event.photogrammetry.focal_length_type, imaging_event.photogrammetry.background_removal, imaging_event.photography.lens_make, imaging_event.photography.lens_model, imaging_event.photography.light_source, processing_event.creator, processing_event.date_created, processing_event.software, processing_event.description)
    
    
    #Append row to dataframe Matrix1
    Matrix1<-rbind(Matrix1,Dataline)
  }
}
  



###
######
### If scanning on a Bruker system with log metadata files
######
###


if(Scanning.system=="Bruker"){
  
  #get the names and number of files in DIRECTORY for loop
  list.files(path = DIRECTORY,pattern = ".log")-> Loglist
  length(Loglist)->metadatafileNo
  #make a blank data frame to be filled in by loop
  Matrix1<-data.frame()
  for (i in 1:metadatafileNo){
    
    #set up text reporting for loop
    Sys.sleep(0.1)
    round(i/metadatafileNo*100,digits=2)->P
    
    #Get file in DIRECTORY
    Loglist[i]->DWCTrip
    #report which file is being processed
    message(paste(P,"%","parsing",DWCTrip, sep=" "))
    #Change .xtekct to .zip for media file
    DWCTripZip<-gsub(".log",".zip",as.character(DWCTrip))
    #remove .log and split remaining file name by -, _ or space
    DWCTrip<-gsub(".log","",as.character(DWCTrip))
    as.list(unlist(strsplit(DWCTrip, split=Delimiter1)))->DWCTList
    if(Name.system=="3Part"){
      #return media part and accession number
      AccessionNo<-unlist(DWCTList[3])
      AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
      if(nchar(AddAccessionFirst) >=1){
        AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
      if(nchar(AddAccessionLast) >=1){
        AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
      #return 4th component as media part (if it exists)
      if(length(unlist(DWCTList))>3)
      {
        media.part<-unlist(DWCTList[4])
      }else{
        media.part<-"entire specimen"
      }
    }
    if(Name.system=="2Part"){
      AccessionNo<-unlist(DWCTList[2])
      AccessionNo<-gsub("[^0-9.-]", "", AccessionNo)
      if(nchar(AddAccessionFirst) >=1){
        AccessionNo<-paste(AddAccessionFirst,AccessionNo,sep = "")}
      if(nchar(AddAccessionLast) >=1){
        AccessionNo<-paste(AccessionNo,AddAccessionLast,sep = "")}
      #return 4th component as media part (if it exists)
      if(length(unlist(DWCTList))>2)
      {
        media.part<-unlist(DWCTList[3])
        
      }else{
        
        media.part<-"entire specimen"
      }
    }
    #ping idigbio for Institution, collection and accessionno
    Search1 <- list("scientificname"=list("type"="exists"), 
                    "catalognumber"=AccessionNo,"institutioncode"=Institution,"collectioncode"=CollectionCode)
    ridigbio::idig_search_records(rq = Search1,limit = 11)->results1
    #Pull out relevant info from iDigBio if found
    if(nrow(results1)>=1){
      UUID<-results1$uuid
      occurrenceID<-results1$occurrenceid
      ScientificName<-results1$scientificname
      Genus<-results1$genus
      Genus<-str_to_title(Genus)
      SpeciesName<-gsub(results1$genus,"",as.character(results1$scientificname))
      DateCollected<-results1$`data.dwc:eventDate`
      Collector<-results1$collector
      Lattitude<-results1$geopoint.lat
      Longitude<-results1$geopoint.lon
      
      #report that the search didnt find anything on idigbio
    }else{
      UUID<-""
      occurrenceID<-"Not on IdigBio"
      ScientificName<-"Not on IdigBio"
      Genus<-"Not on IdigBio"
      SpeciesName<-"Not on IdigBio"
      DateCollected<-"Not on IdigBio"
      Collector<-"Not on IdigBio"
      Lattitude<-"Not on IdigBio"
      Longitude<-"Not on IdigBio"
    }
    #open log file and read data into relevant objects
    paste(DIRECTORY,Loglist[i],sep = "/")->Log.path
    conn<-file(Log.path,open="r")
    linn<-readLines(conn)
    close(conn)
    
    linn[grep('^Pixel Size (um)*', linn)]->Voxelsize
    Voxelsize<-readr::parse_number(Voxelsize)
    
    linn[grep('^Number Of Files*', linn)]->ImageNo
    ImageNo<-readr::parse_number(ImageNo)
    
    linn[grep('^Source Voltage (kV)*', linn)]->Voltage
    Voltage<-readr::parse_number(Voltage)
    
    linn[grep('^Source Current', linn)]->Current
    Current<-readr::parse_number(Current)
    
    linn[grep('^Filter=*', linn)]->Filter
    Filter<-gsub("Filter=","",as.character(Filter[1]))
    
    linn[grep('^Frame Averaging=ON =*', linn)]->Avg
    Avg<-readr::parse_number(Avg)

    
    linn[grep('^Exposure*', linn)]->TimingVal
    TimingVal<-readr::parse_number(TimingVal)
    TimingVal<-TimingVal/1000
    
    linn[grep('^Camera to Source*', linn)]->FDD
    FDD<-readr::parse_number(FDD)
    
    linn[grep('^Object to Source*', linn)]->FOD
    FOD<-readr::parse_number(FOD)
    
    linn[grep('^Camera Pixel Size*', linn)]->PixelsizeX
    PixelsizeX<-readr::parse_number(PixelsizeX)
    linn[grep('^Number Of Rows= =*', linn)]->NrPixelsX
    NrPixelsX<-readr::parse_number(NrPixelsX)
    linn[grep('^Camera Pixel Size*', linn)]->PixelsizeY
    PixelsizeY<-readr::parse_number(PixelsizeY)
    linn[grep('^Number Of Columns*', linn)]->NrPixelsY
    NrPixelsY<-readr::parse_number(NrPixelsY)
    (Voltage*Current/1000000)->Wattage
    
    #bind all data into a row
    Dataline <-cbind(DWCTripZip, media.preview_file, media_type, media_raw_or_derived, media.parent_file, media.parent_ms_id, biological_specimen.ms_id, occurrenceID, biological_specimen.occurrence_id, Institution, CollectionCode, AccessionNo, media.part, media.short_description, media.side, media.description, Creator, media.orientation, media.identifier, media.keyword, media.date_created, media.related_url, Voxelsize, Voxelsize, Voxelsize, Voxelsize, media.series_type, media.unit, media.map_type, biological_specimen.identifier, biological_specimen.related_url, DateCollected, Collector, biological_specimen.description, Lattitude, Longitude, biological_specimen.numeric_time, biological_specimen.original_location, biological_specimen.periodic_time, biological_specimen.is_type_specimen, biological_specimen.sex, biological_specimen.vouchered, Genus, SpeciesName, taxonomy.taxonomy_subspecies, imaging_event.description, Creator, imaging_event.software, imaging_event.date_created, TimingVal, imaging_event.ct.flux_normalization, imaging_event.ct.geometric_calibration, imaging_event.ct.shading_correction, Filter, Avg, ImageNo, Voltage, Wattage, Current, imaging_event.ct.surrounding_material, imaging_event.ct.xray_tube_type, imaging_event.ct.target_type, imaging_event.ct.detector_type, NrPixelsX, PixelsizeX, NrPixelsY, PixelsizeY, imaging_event.ct.detector_configuration, FOD, FDD, imaging_event.ct.target_material, imaging_event.ct.rotation_number, imaging_event.ct.phase_contrast, imaging_event.ct.optical_magnification, imaging_event.ct.acquisition_type, imaging_event.photogrammetry.focal_length_type, imaging_event.photogrammetry.background_removal, imaging_event.photography.lens_make, imaging_event.photography.lens_model, imaging_event.photography.light_source, processing_event.creator, processing_event.date_created, processing_event.software, processing_event.description)
    
    
    #Append row to dataframe Matrix1
    Matrix1<-rbind(Matrix1,Dataline)
  }
}
  
  #rename matrix1 collumns 
colnames(Matrix1)<-c("media.media_file","media.preview_file","media.media_type", "media.raw_or_derived", "media.parent_file", "media.parent_ms_id", "biological_specimen.ms_id", "biological_specimen.idigbio_uuid", "biological_specimen.occurrence_id", "biological_specimen.institution_code", "biological_specimen.collection_code", "biological_specimen.catalog_number", "media.part", "media.short_description", "media.side", "media.description", "media.creator", "media.orientation", "media.identifier", "media.keyword", "media.date_created", "media.related_url", "media.x_spacing", "media.y_spacing", "media.z_spacing", "media.slice_thickness", "media.series_type", "media.unit", "media.map_type", "biological_specimen.identifier", "biological_specimen.related_url", "biological_specimen.date_created", "biological_specimen.creator", "biological_specimen.description", "biological_specimen.latitude", "biological_specimen.longitude", "biological_specimen.numeric_time", "biological_specimen.original_location", "biological_specimen.periodic_time", "biological_specimen.is_type_specimen", "biological_specimen.sex", "biological_specimen.vouchered", "taxonomy.taxonomy_genus", "taxonomy.taxonomy_species", "taxonomy.taxonomy_subspecies", "imaging_event.description", "imaging_event.creator", "imaging_event.software", "imaging_event.date_created", "imaging_event.ct.exposure_time", "imaging_event.ct.flux_normalization", "imaging_event.ct.pixel_spacing_calibration", "imaging_event.ct.shading_correction", "imaging_event.ct.ie_filter", "imaging_event.ct.frame_averaging", "imaging_event.ct.projections", "imaging_event.ct.voltage", "imaging_event.ct.power", "imaging_event.ct.amperage", "imaging_event.ct.surrounding_material", "imaging_event.ct.xray_tube_type", "imaging_event.ct.target_type", "imaging_event.ct.detector_type", "imaging_event.ct.detector_pixels_x", "imaging_event.ct.detector_pixel_size_x", "imaging_event.ct.detector_pixels_y", "imaging_event.ct.detector_pixel_size_y", "imaging_event.ct.detector_configuration", "imaging_event.ct.source_object_distance", "imaging_event.ct.source_detector_distance", "imaging_event.ct.target_material", "imaging_event.ct.rotation_number", "imaging_event.ct.phase_contrast", "imaging_event.ct.optical_magnification", "imaging_event.ct.acquisition_type", "imaging_event.photogrammetry.focal_length_type", "imaging_event.photogrammetry.background_removal", "imaging_event.photography.lens_make", "imaging_event.photography.lens_model", "imaging_event.photography.light_source", "processing_event.creator", "processing_event.date_created", "processing_event.software","processing_event.description")  
sub("NA","",paste(Matrix1$biological_specimen.latitude))->Matrix1$biological_specimen.latitude
sub("NA","",paste(Matrix1$biological_specimen.longitude))->Matrix1$biological_specimen.longitude
sub("T00:00:00.00:00","",Matrix1$biological_specimen.date_created)->Matrix1$biological_specimen.date_created
sub("NA","",paste(Matrix1$biological_specimen.date_created))->Matrix1$biological_specimen.date_created
cbind("","",Matrix1)->Matrix1

#go to DIRECTORY, save CSV and return to original directory
setwd(DIRECTORY)
gsub("/","",as.character(DIRECTORY))->DIRFolder
paste(DIRFolder,"morphosource batch.xlsx",sep = " ")->fileName

write.xlsx(Matrix1,file = fileName,row.names = FALSE)
setwd(mypath)
return(Matrix1)
}

