/*

Check last sync date

Table: [sqlDDMM].[dbo].[Settings]

Query file: Settings-PrjLastSync.sql

*/

/****** Searching for Last Sync dates for specific Projec in Settings ******/
DECLARE /* PROJECT CODE (Use % not to filter on specific Project) */
@PrjCode VARCHAR(MAX) = '4300_U4';

SELECT *
FROM [sqlDDMM].[dbo].[Settings]
WHERE [Project] LIKE @PrjCode



/*

NOTES:

- Time is 2h behind.

- If there's need to let sync process run again on a specific date, only change the day or month (not time) with the following command:

    UPDATE [sqlDDMM].[dbo].[Settings]
    WHERE [Project] = @PrjCode AND [Key] LIKE '%Import%'
    SET [Value] = '2023-05-18 15:09:59.975'

*/