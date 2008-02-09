VIPER (Visual Inspection of Peak/Elution Relationships)

VIPER can be used to visualize and characterize the features detected during
LC-MS analyses.  It is primarily intended for processing deisotoped data from 
high mass measurement accuracy instruments (e.g. FT, TOF, or Orbitrap) and 
comparing the features found to a list of expected peptides (aka mass and 
time (MT) tags), as practiced by the AMT Tag approach (see Zimmer, Monroe, 
Qian, & Smith, Mass Spec Reviews, 2006, 25, 450-482).

The software allows one to display the data as a two-dimensional plot of 
spectrum number (i.e. elution time) vs. mass.  It can read MS data from 
several file formats: .PEK, .CSV, .mzXML, and .mzData.  See below for 
additional details on the various file formats. VIPER can process a data 
file in an automated fashion to:
 1)	Load and filter the data
 2)	Find features and their chromatographic peak profile
 3)	Align the observed elution times of the features with the elution times
    of the expected peptides (i.e. MT Tags)
 4)	Refine the mass calibration
 5)	Match the features to the MT tags
 6)	Export the results to a report file

VIPER has a rich, full featured graphical user interface with numerous options 
for exploring the data and customizing the analysis parameters.  It runs on 
Windows computers and requires Microsoft Access be installed to create and edit
the MT tag databases.  In order to view mass spectra from raw data files (e.g.
Finnigan .Raw files, you will also need to install ICR-2LS (see the .PEK files 
section below for more info).

Double click the VIPER_Installer.msi file to install.  The application will be 
installed, along with the LCMSFeatureFinder.exe program (and required DLLs).  
Note that the LCMSFeatureFinder requires that the Microsoft .NET Framework v1.1 
be installed. See http://msdn2.microsoft.com/en-us/netframework/aa569264.aspx 
for instructions on how to validate that the Framework is installed.  Follow 
this link to install it: http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=262D25E3-F589-4842-8157-034D1E7CF3A3

If, after installing VIPER, the LCMSFeatureFinder does not run properly, then 
install it separately using the LCMSFeatureFinder.msi file (aka "LCMSFeatureFinder (Install this after installing Viper).msi")

The shortcut for starting VIPER can be found at Start Menu -> Programs -> PAST Toolkit -> VIPER


Microsoft Access MT Database file formats:
	- An Access database is used to store the details of the MT tags to match against
	- You can have multiple Microsoft Access database files, each with a different
	  set of MT tags.  Use "Steps->3. Select MT Tags (Connect to DB)" to specify
	  or change the database.
	- There are two primary Access databas file formats:
	
	- ** AMT Table File Format (table AMT contains all the key information) **
	- The Access database file must have a table named AMT, containing the fields:
		 AMT_ID
		 AMTMonoisotopicMass
		 NET
		 PNET or RetentionTime
	  Note that the RetentionTime field is skipped if PNET is present 
	- The AMT table can also optionally contain the fields: 
		MSMS_Obs_Count
		High_Normalized_Score
		High_Discriminant_Score
		NitrogenAtom
		AA_Cystine_Count
		Status
		Peptide
	- The Access database file can optionally contain two tables containing Protein
	  information and MT to Protein Mapping information.  If these tables are
	  present, then Viper can optionally display the protein names along with
	  the MT tags matched.
		- The Protein table must be named Proteins, containing the fields:
			Protein_ID
			Protein_Name
		- The MT to Protein Mapping table must be named AMT_to_Protein_Map, with fields:
			AMT_ID
			Protein_ID
		- The advantage of using two tables to track protein information is
		  that additional information for each protein can be defined in the 
		  Proteins table while the AMT_to_Protein_Map simply associates MT tags
		  with proteins
	- File QCStandards.mdb demonstrates this database format
		(installed by VIPER_Installer.msi at C:\Program Files\Viper\)

	- ** Expanded File Format (key information stored in multiple tables) **
	- The Access database file must have two tables, T_Mass_Tags and T_Mass_Tags_NET
	- T_Mass_Tags must have columns Mass_Tag_ID and Monoisotopic_Mass
		- Other typical columns include Peptide, Peptide_Obs_Count_Passing_Filter, 
		  High_Normalized_Score, High_Discriminant_Score, and High_Peptide_Prophet_Probability
	- T_Mass_Tags_NET must have columns Mass_Tag_ID and Avg_GANET
		- Other typical columns include Cnt_GANET, StD_GANET, and PNET
	- The Access database file can optionally contain two tables containing Protein
	  information and MT to Protein Mapping information.  If these tables are
	  present, then Viper can optionally display the protein names along with
	  the MT tags matched.
		- The Protein table must be named T_Proteins, containing the fields:
			Ref_ID
			Reference
		- The MT to Protein Mapping table must be named T_Mass_Tag_to_Protein_Map, with fields:
			Mass_Tag_ID
			Ref_ID
		- The advantage of using two tables to track protein information is
		  that additional information for each protein can be defined in the 
		  Proteins table while the T_Mass_Tag_to_Protein_Map simply associates 
		  MT tags with proteins
	- File QCStandards_ExpandedFormat.mdb demonstrates this database format
		(installed by VIPER_Installer.msi at C:\Program Files\Viper\)

	- The database can contain other tables, but Viper will only read data from 
	  the key tables.  Also, the tables can contain other columns besides those 
	  mentioned above, but Viper will only read the data from the known columns.
	- The column order in the tables does not matter; Viper can cope with rearranged columns.


Auto-analysis files:
	- You can save a .Ini file using File->Save/Load/Edit Analysis Settings to 
	  capture all of the parameters used for analyzing data.  This .Ini file 
	  can be used to auto-analyze one or more input files.  For more information,
	  please run VIPER from the command line (aka the command prompt or Dos shell)
	  using "VIPER /?" to see a help screen with info on initiating automated analyses.


Input file format notes:
.PEK files, generated by ICR-2LS (download at http://ncrr.pnl.gov/software/)
	- ICR-2LS processes mass spectrometry data from several different vendor's 
	  mass spectrometers to deisotope (deconvolute) each spectrum, determining
	  the monoisotopic mass and charge of the compounds in each spectrum and
	  storing the results in a text file called a .PEK file
	- PEK files are text files with a multi-line header region and tab delimited 
	  data section for each mass spectrum.  Separate sections are included for 
	  charge state (un-deisotoped) and isotopic (deisotoped) data; VIPER can
	  read and process both types of data.
	- For isotopic data, each spectrum in a .Pek file must, at a minimum, 
	  match this format:

		Filename: MyDataset.pek.00256
		CS,  Abundance,   m/z,   Fit,    Average MW, Monoisotopic MW,    Most abundant MW
		2	438000	574.243226	0.026	1147.171	1146.4719	1146.4719
		3	210000	534.554343	0.115	1601.7039	1600.6412	1600.6414
		3	299000	778.352243	0.069	2333.5481	2332.0349	2333.0371
		2	620000	566.245476	0.055	1131.1636	1130.4764	1130.4762
		3	388000	795.35626	0.094	2384.5898	2383.04695	2384.0505

		- The scan number of the spectrum is defined by the numbers at the end
		  of the "Filename:" line.  In the above example, the scan number is 256
		- The "CS,  Abundance," line indicates that the data directly after it is
		  deisotoped data.  Note that "CS," and "Abundance" must be separated 
		  by two spaces.
		- The data lines are tab delimited lists of the deisotoped data in a spectrum
	- For charge state data, each spectrum in a .Pek file must, at a minimum, 
	  match this format:

		Filename: MyDataset.pek.00256
		First CS,    Number of CS,   Abundance,   Mass,   Standard deviation
		13	3	3.86E+04	9439.38	.6473	
		3	3	1.71E+04	3143.49	.9077
	
		- Again, the scan number of the spectrum is defined by the numbers at
		  the end of the "Filename:" line.
		- The "First CS," line indicates that the data directly after it is
		  un-deisotoped data
	- An example PEK file is installed by VIPER_Installer.msi

.CSV files, generated by Decon2LS (available at http://ncrr.pnl.gov/software/)
	- Decon2LS is second generation software that builds upon the algorithms 
	  present in ICR-2LS, and thus also capable of deisotoping mass spectra.
	  Decon2LS saves its results in a pair of CSV (comma-separated value) files
	- The _scans.csv file contains information about each mass spectrum (aka scan); columns are:
		scan_num
		scan_time
		type
		bpi
		bpi_mz
		tic
		num_peaks
		num_deisotoped
	- The _isos.csv file contains the deisotoped data; columns are:
		scan_num
		charge
		abundance
		mz
		fit
		average_mw
		monoisotopic_mw
		mostabundant_mw
		fwhm
		signal_noise
		mono_abundance
		mono_plus2_abundance
	- Example CSV files are installed by VIPER_Installer.msi

.mzXML files are XML files that conform to the mzXML MS data file standard
	- For more information, see http://sashimi.sourceforge.net/software_glossolalia.html
	- mzXML file names must end with either .mzXML or _mzXML.xml
	- An example mzXML file is installed by VIPER_Installer.msi

.mzData files are XML files that conform to the mzData MS data file standard
	- For more information, see http://psidev.sourceforge.net/ms/
	- mzData file names must end with either .mzData or _mzData.xml
	- An example mzData file is installed by VIPER_Installer.msi


Output file format notes:
	- Viper saves a binary file with extension .Gel containing all of the data
	  loaded, the LC-MS Features found, the PMT tags matched, and all the options used.  
	  This is a fairly large file with arrays of user defined types (UDTs) saved 
	  in binary format, and is thus a bit tricky to parse.  Contact the author if 
	  you'd like to receive more information.

	- Viper saves tab-delimited text files for most of the exporting options.
	  These files will have a header line with tab-delimited headers, then 
	  tab-delimited data lines.


-------------------------------------------------------------------------------
Written by Matthew Monroe and Nikola Tolic for the Department of Energy (PNNL, Richland, WA)
Copyright 2006, Battelle Memorial Institute.  All Rights Reserved.

E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

All publications that result from the use of this software should include 
the following acknowledgment statement:
 Portions of this research were supported by the U.S. Department of Energy 
 Office of Biological and Environmental Research Genomes:GtL Program, the NIH 
 National Center for Research Resources (Grant RR018522), and the National 
 Institute of Allergy and Infectious Diseases (NIH/DHHS through interagency 
 agreement Y1-AI-4894-01).  PNNL is operated by Battelle Memorial Institute 
 for the U.S. Department of Energy under contract DE-AC05-76RL0 1830.

Notice: This computer software was prepared by Battelle Memorial Institute, 
hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
Department of Energy (DOE).  All rights in the computer software are reserved 
by DOE on behalf of the United States Government and the Contractor as 
provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
SOFTWARE.  This notice including this sentence must appear on any copies of 
this computer software.
