-- Use this to populate the AMT table
SELECT DISTINCT 
    MT.Mass_Tag_ID AS AMT_ID, 
    MT.Monoisotopic_Mass AS AMTMonoisotopicMass, 
    MTN.Avg_GANET AS NET, MTN.PNET, 
    MT.Peptide_Obs_Count_Passing_Filter AS MSMS_Obs_Count, 
    MT.High_Normalized_Score, 
    MT.High_Discriminant_Score
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tags_NET MTN ON 
    MT.Mass_Tag_ID = MTN.Mass_Tag_ID INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins P ON MTPM.Ref_ID = P.Ref_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (P.Reference LIKE 'p0%' OR
    P.Reference LIKE '0%') AND (P.Reference NOT LIKE 'p011%')
ORDER BY MT.Mass_Tag_ID

-- Use this to populate the AMT_Proteins table
SELECT DISTINCT 
    P.Ref_ID AS Protein_ID, P.Reference AS Protein_Name
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins P ON MTPM.Ref_ID = P.Ref_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (P.Reference LIKE 'p0%' OR
    P.Reference LIKE '0%') AND (P.Reference NOT LIKE 'p011%')
ORDER BY P.Protein_ID

-- Use this to populate the AMT_to_Protein_Map table
SELECT MTPM.Mass_Tag_ID, MTPM.Ref_ID
FROM T_Mass_Tags MT INNER JOIN
    T_Mass_Tag_to_Protein_Map MTPM ON 
    MT.Mass_Tag_ID = MTPM.Mass_Tag_ID INNER JOIN
    T_Proteins P ON MTPM.Ref_ID = P.Ref_ID
WHERE (MT.PMT_Quality_Score >= 2) AND 
    (MT.Peptide_Obs_Count_Passing_Filter >= 5) AND 
    (P.Reference LIKE 'p0%' OR
    P.Reference LIKE '0%') AND (P.Reference NOT LIKE 'p011%')
ORDER BY MT.Mass_Tag_ID, MTPM.Ref_ID
