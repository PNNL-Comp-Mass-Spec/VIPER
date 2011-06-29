-- Connect to server Porky
USE MT_Shewanella_ProdTest_Formic_P460

Declare @PMTQualityScoreFilter int
Declare @PeptideObsCountFilter int

Set @PMTQualityScoreFilter = 2
Set @PeptideObsCountFilter = 3

-- Data for T_Mass_Tags
SELECT Mass_Tag_ID, Peptide, Monoisotopic_Mass, Multiple_Proteins, Created, Last_Affected, Number_Of_Peptides, 
    Peptide_Obs_Count_Passing_Filter, High_Normalized_Score, High_Discriminant_Score, High_Peptide_Prophet_Probability, 
    Min_Log_EValue, Mod_Count, Mod_Description, PMT_Quality_Score
FROM T_Mass_Tags
WHERE (PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
ORDER BY Mass_Tag_ID

-- Data for T_Mass_Tags_NET
SELECT MT.Mass_Tag_ID, MTN.Min_GANET, MTN.Max_GANET, 
       MTN.Avg_GANET, MTN.Cnt_GANET, 
       ISNULL(MTN.StD_GANET, 0) AS StD_GANET, 
       ISNULL(MTN.StdError_GANET, 0) AS StdError_GANET, 
       MTN.PNET
FROM T_Mass_Tags MT
     INNER JOIN T_Mass_Tags_NET MTN
       ON MT.Mass_tag_ID = MTN.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND NOT
      AVG_GANET Is Null
ORDER BY Mass_Tag_ID


-- Populate T_Proteins
SELECT Prot.Ref_ID, Prot.Reference, Prot.Description, 
    Prot.Protein_Residue_Count, Prot.Monoisotopic_Mass, 
    Prot.Protein_Collection_ID, Prot.Last_Affected,
    '' AS Protein_Sequence, Prot.Protein_DB_ID, 
    Prot.External_Reference_ID, Prot.External_Protein_ID
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins Prot ON MTPM.Ref_ID = Prot.Ref_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
GROUP BY Prot.Ref_ID, Prot.Reference, Prot.Description, 
    Prot.Protein_Residue_Count, Prot.Monoisotopic_Mass, 
    Prot.Protein_Collection_ID, Prot.Last_Affected,
    Prot.Protein_DB_ID, 
    Prot.External_Reference_ID, Prot.External_Protein_ID
ORDER BY Prot.Ref_ID

-- Populate T_Proteins with proteins that have a Null Protein_Collection_ID
SELECT Prot.Ref_ID, Prot.Reference, Prot.Reference AS Description, 
    0 AS Protein_Residue_Count, 0 AS Monoisotopic_Mass, 
    0 AS Protein_Collection_ID, Prot.Last_Affected,
    '' AS Protein_Sequence, 0 AS Protein_DB_ID, 
    0 AS External_Reference_ID, 0 AS External_Protein_ID
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins Prot ON MTPM.Ref_ID = Prot.Ref_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
      AND Protein_Collection_ID Is Null
GROUP BY Prot.Ref_ID, Prot.Reference, Prot.Description, 
    Prot.Protein_Residue_Count, Prot.Monoisotopic_Mass, 
    Prot.Protein_Collection_ID, Prot.Last_Affected,
    Prot.Protein_DB_ID, 
    Prot.External_Reference_ID, Prot.External_Protein_ID
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
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
ORDER BY MT.Mass_Tag_ID, MTPM.Ref_ID

-- Populate T_Mass_Tag_to_Protein_Map with entries where Mass_Tag_Name is null
SELECT MTPM.Mass_Tag_ID, '' AS Mass_Tag_Name, 
    MTPM.Ref_ID, MTPM.Cleavage_State, 
    0 AS Fragment_Number, 0 AS Fragment_Span, 
    0 AS Residue_Start, 0 AS Residue_End, 
    0 AS Repeat_Count, MTPM.Terminus_State, 
    0 AS Missed_Cleavage_Count
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND
      Mass_Tag_Name IS NULL
ORDER BY MT.Mass_Tag_ID, MTPM.Ref_ID


-- Populate T_Mass_Tag_Conformers_Observed
SELECT MTC.Conformer_ID, MTC.Mass_Tag_ID, MTC.Charge, MTC.Conformer, 
       MTC.Drift_Time_Avg, 
       ISNULL(MTC.Drift_Time_StDev, 0) AS Drift_Time_StDev, 
       MTC.Obs_Count, MTC.Last_Affected
FROM T_Mass_Tags MT INNER JOIN
     T_Mass_Tag_Conformers_Observed MTC ON 
     MT.Mass_Tag_ID = MTC.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
ORDER BY MTC.Conformer_ID


-- No longer used: T_Mass_Tag_Peptide_Prophet_Stats
SELECT MTP.Mass_Tag_ID, MTP.ObsCount_CS1, MTP.ObsCount_CS2, MTP.ObsCount_CS3, MTP.PepProphet_FScore_Avg_CS1, 
    MTP.PepProphet_FScore_Avg_CS2, MTP.PepProphet_FScore_Avg_CS3
FROM T_Mass_Tag_Peptide_Prophet_Stats MTP INNER JOIN
     T_Mass_Tags MT ON MTP.Mass_Tag_ID = MT.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
      (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
ORDER BY MTP.Mass_Tag_ID



-- The following tables do not need to be populated

-- Populate V_Filter_Set_Overview_Ex
SELECT *
FROM V_Filter_Set_Overview_Ex

-- Find jobs to use
SELECT *
FROM ( SELECT *,
              ROW_NUMBER() OVER ( PARTITION BY Instrument ORDER BY MTCount DESC ) AS InstrumentRank
       FROM ( SELECT AD.Job,
                     COUNT(DISTINCT MT.Mass_Tag_ID) AS MTCount,
                     AD.Dataset,
                     AD.Dataset_Created_DMS,
                     AD.Instrument,
                     AD.Analysis_Tool
              FROM T_Mass_Tags MT
                   INNER JOIN T_Peptides Pep
                     ON MT.Mass_Tag_ID = Pep.Mass_Tag_ID
                   INNER JOIN T_Analysis_Description AD
                     ON Pep.Analysis_ID = AD.Job
              WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND
                    (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter)
              GROUP BY AD.Job, AD.Dataset, AD.Dataset_Created_DMS, AD.Instrument, AD.Analysis_Tool ) 
              LookupQ ) InstrumentQ
WHERE InstrumentRank <= 2


-- Populate T_Analysis_Description using the desired jobs
SELECT Job, Dataset, Dataset_ID, Dataset_Created_DMS, 
    Dataset_Acq_Time_Start, Dataset_Acq_Time_End, 
    Dataset_Scan_Count, Experiment, Campaign, Experiment_Organism, 
    Instrument_Class, Instrument, Analysis_Tool, 
    Parameter_File_Name, Settings_File_Name, 
    Organism_DB_Name, Protein_Collection_List, 
    Protein_Options_List, Completed, ResultType, 
    Separation_Sys_Type, ScanTime_NET_Slope, 
    ScanTime_NET_Intercept, ScanTime_NET_RSquared, 
    ScanTime_NET_Fit
FROM T_Analysis_Description
WHERE (Job IN (287491,317820,312973,317831,312977,263718,287343,265349))

-- Populate T_Peptides using the desired jobs
SELECT Pep.Peptide_ID, Pep.Analysis_ID, Pep.Scan_Number, 
    Pep.Number_Of_Scans, Pep.Charge_State, Pep.MH, 
    Pep.Multiple_Proteins, Pep.Peptide, Pep.Mass_Tag_ID, 
    Pep.GANET_Obs, Pep.Scan_Time_Peak_Apex, Pep.Peak_Area, 
    Pep.Peak_SN_Ratio
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND 
    (Pep.Analysis_ID IN (287491,317820,312973,317831,312977,263718,287343,265349))
ORDER BY Pep.Analysis_ID, MT.Mass_Tag_ID

-- Populate T_Score_Discriminant using the desired jobs
-- The paste append operation can take a long time because Access is checking the foreign key relationship with T_Peptides
SELECT Pep.Peptide_ID, 
       ISNULL(SD.Peptide_Prophet_FScore, 0) AS Peptide_Prophet_FScore, 
       ISNULL(SD.Peptide_Prophet_Probability, 0) AS Peptide_Prophet_Probability, 
       ISNULL(SD.MSGF_SpecProb, 1) AS MSGF_SpecProb
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Score_Discriminant SD ON 
    Pep.Peptide_ID = SD.Peptide_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND 
    (Pep.Analysis_ID IN (287491,317820,312973,317831,312977,263718,287343,265349))
ORDER BY Pep.Peptide_ID

-- Populate T_Score_Sequest using the desired jobs
-- The paste append operation can take a long time because Access is checking the foreign key relationship with T_Peptides
SELECT Pep.Peptide_ID, SS.XCorr, SS.DeltaCn2, SS.Sp, SS.RankXc, SS.DelM
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Score_Sequest SS ON 
    Pep.Peptide_ID = SS.Peptide_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND 
    (Pep.Analysis_ID IN (287491,317820,312973,317831,312977,263718,287343,265349))
ORDER BY Pep.Peptide_ID

-- Populate T_Score_XTandem using the desired jobs
SELECT Pep.Peptide_ID, 
		X.Hyperscore,
		X.Log_EValue,
		X.DeltaCn2,
		X.Y_Score,
		X.Y_Ions,
		X.B_Score,
		X.B_Ions,
		X.DelM,
		X.Intensity,
		X.Normalized_Score
FROM T_Mass_Tags MT INNER JOIN
    T_Peptides Pep ON 
    MT.Mass_Tag_ID = Pep.Mass_Tag_ID INNER JOIN
    T_Score_XTandem X ON 
    Pep.Peptide_ID = X.Peptide_ID
WHERE (MT.PMT_Quality_Score >= @PMTQualityScoreFilter) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= @PeptideObsCountFilter) AND 
    (Pep.Analysis_ID IN (287491,317820,312973,317831,312977,263718,287343,265349))
ORDER BY Pep.Peptide_ID


