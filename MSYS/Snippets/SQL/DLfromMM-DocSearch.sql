/*

Check if documents have been imported from MM to SPO (search for them by TCM Document Number and Client Code)

Table: [sqlDDMM].[dbo].[TB_DLfromMM]

Query file: DLfromMM-DocSearch.sql

*/

/****** Searching for Document by TCM Document Number and Client Code in TB_DLfromMM ******/

/* Search criteria */
DECLARE @SearchByTCM_DNs BIT = 1 -- Set to 1 to search by TCM_DNs, 0 to ignore (No need to clear variables)
DECLARE @SearchByClientCodes BIT = 1 -- Set to 1 to search by ClientCodes, 0 to ignore (No need to clear variables)

/* TCM DOCUMENT NUMBER ARRAY TO BE SEARCHED (Comma separated) */
DECLARE @TCM_DNs TABLE (TCM_DN VARCHAR(MAX)) INSERT INTO @TCM_DNs VALUES
('%K439-VZ-DP-00000PFD001000901%')

/* CLIENT CODE ARRAY TO BE SEARCHED (Comma separated) */
DECLARE @ClientCodes TABLE (ClientCode VARCHAR(MAX)) INSERT INTO @ClientCodes VALUES
('%K484-3900-0000-CN-0001-01%'),
('%K484-3900-0000-CN-0001-02%')

/* Run the query */
SELECT *
FROM [sqlDDMM].[dbo].[TB_DLfromMM]
WHERE
(
    @SearchByTCM_DNs = 1 AND
    EXISTS (
        SELECT 1
        FROM @TCM_DNs
        WHERE [sqlDDMM].[dbo].[TB_DLfromMM].[TCM DOCUMENT CODE] LIKE TCM_DN
    )
)
OR
(
    @SearchByClientCodes = 1 AND
    EXISTS (
        SELECT 1
        FROM @ClientCodes
        WHERE [sqlDDMM].[dbo].[TB_DLfromMM].[CLIENT CODE] LIKE ClientCode
    )
)