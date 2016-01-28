VIPER version of STAC.

-m: AMT database with columns: Mass_Tag_ID, Monoisotopic_Mass, Avg_GANET, High_Peptide_Prophet_Probability, Cnt_GANET
-u: UMC file with columns: UMCIndex, NETClassRep, UMCMonoMW
-odir: Directory to write files to.  Creates files with the same base string as the UMC file and adds the suffixes _STAC.csv and _FDR.csv for the STAC score file and FDR summary file, respectively.
-useP: True/False switch to turn off use of probability scores.  Defaults to false, so must be set to true to use probabilities.
-ppm and -NET: Mass and NET tolerances.  Will work best if these are the same (refined) tolerances used for 'Database Search' step.

Usage: STAC.exe
 -m [MassTagFile]  -u [UMC File] -odir [output directory] [optional arguments]
 -m: Mass Tag File. For performing peak matching this is the file
         with all mass tags.
         Required Columns: Mass_Tag_ID, Monoisotopic_Mass,Avg_GANET,
         [High_Peptide_Prophet_Probability,] Cnt_GANET
 -u: UMC File. Text file with aligned UMCs.  Required columns:
         UMCIndex, NETClassRep, UMCMonoMW
 -odir: Output directory. Directory to write match scores to.

*********OPTIONAL AGUMENTS FOR DETAILED SPECIFICATION***********.

 -useP [T/F] : Whether or not to use optionally provided prior probabilities
 -ppm [value] : Maximum mass tolerance in ppm
 -NET [value] : Maximum NET tolerance