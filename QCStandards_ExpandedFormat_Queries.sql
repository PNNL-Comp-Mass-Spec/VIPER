-- Data for T_Mass_Tags
SELECT Mass_Tag_ID, Peptide, Monoisotopic_Mass, 
    Multiple_Proteins, Created, Last_Affected, 
    Number_Of_Peptides, Peptide_Obs_Count_Passing_Filter, 
    High_Normalized_Score, High_Discriminant_Score, 
    High_Peptide_Prophet_Probability, Min_Log_EValue, 
    Mod_Count, Mod_Description, PMT_Quality_Score
FROM T_Mass_Tags
WHERE (PMT_Quality_Score >= 2) AND 
    (Peptide_Obs_Count_Passing_Filter >= 5)
ORDER BY Mass_Tag_ID



-- Find jobs to use
SELECT AD.Job, COUNT(DISTINCT MT.Mass_Tag_ID) AS MTCount, 
    AD.Dataset, AD.Dataset_Created_DMS, AD.Instrument, 
    AD.Analysis_Tool
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Analysis_Description AD ON 
    Pep.Analysis_ID = AD.Job
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (AD.Dataset LIKE 'qc_05%') AND (AD.Instrument = 'LTQ_3')
GROUP BY AD.Job, AD.Dataset, AD.Dataset_Created_DMS, 
    AD.Instrument, AD.Analysis_Tool

-- Populate T_Analysis_Description
SELECT Job, Dataset, Dataset_ID, Dataset_Created_DMS, 
    Dataset_Acq_Time_Start, Dataset_Acq_Time_End, 
    Dataset_Scan_Count, Experiment, Campaign, Organism, 
    Instrument_Class, Instrument, Analysis_Tool, 
    Parameter_File_Name, Settings_File_Name, 
    Organism_DB_Name, Protein_Collection_List, 
    Protein_Options_List, Completed, ResultType, 
    Separation_Sys_Type, ScanTime_NET_Slope, 
    ScanTime_NET_Intercept, ScanTime_NET_RSquared, 
    ScanTime_NET_Fit
FROM T_Analysis_Description
WHERE (Job IN (109584, 109728, 109730, 109732, 109734))

-- Populate T_Peptides
SELECT Pep.Peptide_ID, Pep.Analysis_ID, Pep.Scan_Number, 
    Pep.Number_Of_Scans, Pep.Charge_State, Pep.MH, 
    Pep.Multiple_Proteins, Pep.Peptide, Pep.Mass_Tag_ID, 
    Pep.GANET_Obs, Pep.Scan_Time_Peak_Apex, Pep.Peak_Area, 
    Pep.Peak_SN_Ratio
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (Pep.Analysis_ID IN (109584, 109728, 109730, 109732, 109734))
ORDER BY Pep.Analysis_ID, MT.Mass_Tag_ID

-- Populate T_Score_Discriminant
SELECT Pep.Peptide_ID, SD.Peptide_Prophet_FScore, 
    SD.Peptide_Prophet_Probability
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Score_Discriminant SD ON 
    Pep.Peptide_ID = SD.Peptide_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (Pep.Analysis_ID IN (109584, 109728, 109730, 109732, 109734))
ORDER BY Pep.Peptide_ID

-- Populate T_Score_Sequest
SELECT Pep.Peptide_ID, SS.XCorr, SS.DeltaCn2, SS.Sp, SS.RankXc, 
    SS.DelM
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Score_Sequest SS ON 
    Pep.Peptide_ID = SS.Peptide_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (Pep.Analysis_ID IN (109584, 109728, 109730, 109732, 109734))
ORDER BY Pep.Peptide_ID


-- Populate T_Proteins
SELECT Prot.Ref_ID, Prot.Reference, Prot.Description, 
    Prot.Protein_Residue_Count, Prot.Monoisotopic_Mass, 
    Prot.Protein_Collection_ID, Prot.Last_Affected
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins Prot ON MTPM.Ref_ID = Prot.Ref_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5)
GROUP BY Prot.Ref_ID, Prot.Reference, Prot.Description, 
    Prot.Protein_Residue_Count, Prot.Monoisotopic_Mass, 
    Prot.Protein_Collection_ID, Prot.Last_Affected
ORDER BY Prot.Ref_ID


-- Populate T_Mass_Tag_to_Protein_Map
SELECT MTPM.Mass_Tag_ID, MTPM.Mass_Tag_Name, 
    MTPM.Ref_ID, MTPM.Cleavage_State, 
    MTPM.Fragment_Number, MTPM.Fragment_Span, 
    MTPM.Residue_Start, MTPM.Residue_End, 
    MTPM.Repeat_Count, MTPM.Terminus_State, 
    MTPM.Missed_Cleavage_Count
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5)
ORDER BY MT.Mass_Tag_ID, MTPM.Ref_ID
