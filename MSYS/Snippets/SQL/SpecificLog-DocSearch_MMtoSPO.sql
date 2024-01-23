/*
Specific logs

Table: [sqlDDMM].[dbo].[LOG_TB_DLfromMM]

Query file: SpecificLog-DocSearch.sql

*/

/****** Searching for specific Logs that contain Document by TCM Document Number or Client Code in LOG_TB_DLfromMM ******/

/* Search criteria  (Set both to 0 to search for the whole PROJECT) */
DECLARE @SearchByTCM_DNs BIT = 1 -- Set to 1 to search by TCM_DNs, 0 to ignore (No need to clear variables)
DECLARE @SearchByClientCodes BIT = 0 -- Set to 1 to search by ClientCodes, 0 to ignore (No need to clear variables)

DECLARE /* PROJECT CODE */
@PrjCode VARCHAR(MAX) = 'K439';

/* TCM DOCUMENT NUMBER ARRAY TO BE SEARCHED (Comma separated) */
DECLARE @TCM_DNs TABLE (TCM_DN VARCHAR(MAX)) INSERT INTO @TCM_DNs VALUES
('%K439-VZ-DP-00000PFD001000901%')

/* CLIENT CODE ARRAY TO BE SEARCHED (Comma separated) */
DECLARE @ClientCodes TABLE (ClientCode VARCHAR(MAX)) INSERT INTO @ClientCodes VALUES
('%U4-LG-316-00489%')

/* Run the query */
SELECT *
FROM [sqlDDMM].[dbo].[LOG_TB_DLfromMM]
WHERE
[PJT_CODE] LIKE @PrjCode AND
(
    (
        @SearchByTCM_DNs = 0 AND
        @SearchByClientCodes = 0
    )
    OR
    (
        (
            @SearchByTCM_DNs = 1 AND
            EXISTS (
                SELECT 1
                FROM @TCM_DNs
                WHERE [sqlDDMM].[dbo].[LOG_TB_DLfromMM].[TCM DOCUMENT CODE] LIKE TCM_DN
            )
        )
        OR
        (
            @SearchByClientCodes = 1 AND
            EXISTS (
                SELECT 1
                FROM @ClientCodes
                WHERE [sqlDDMM].[dbo].[LOG_TB_DLfromMM].[CLIENT CODE] LIKE ClientCode
            )
        )
    )
)
ORDER BY [DATE_LOG] DESC;