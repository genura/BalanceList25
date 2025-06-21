/*
SQL server 2008 uyumlu !
programming by Cuneyt YENER

quanticvision.co.uk

*/

WITH LastPaymentCTE AS (
    SELECT
        A1.SYYY,
        CASE 
            WHEN ISNUMERIC(A1.TDATE1) = 1 THEN CONVERT(datetime, CONVERT(varchar(8), A1.TDATE1))
            ELSE NULL
        END AS LastPaymentDate,
        A1.CREDIT AS LastPaymentAmount,
        A1.SOURCE AS LastPaymentType,
        CL.CTYPE,
        ROW_NUMBER() OVER (PARTITION BY A1.SYYY ORDER BY A1.TDATE1 DESC, A1.CREDIT DESC) AS rn
    FROM dbo.ACC_TRANSACTION A1
    LEFT JOIN dbo.COLLECTION CL ON A1.SYYY = CL.SYYY
    WHERE A1.CREDIT > 0
),
LastInvoiceCTE AS (
    SELECT
        SYYY,
        MAX(
            CASE 
                WHEN ISNUMERIC(INVOICE_DATE1) = 1 THEN CONVERT(datetime, CONVERT(varchar(8), INVOICE_DATE1))
                ELSE NULL
            END
        ) AS LastInvoiceDate
    FROM dbo.INVOICE_MASTER 	where ITYPE = 1
    GROUP BY SYYY
),
LastInvoiceDetailCTE AS (

SELECT
    IM.SYYY,
    IM.INVOICE_NO,
    IM.MASTERNO,
	OM.ORDER_NO,

    CASE 
        WHEN ISNUMERIC(IM.INVOICE_DATE1) = 1 THEN CONVERT(datetime, CONVERT(varchar(8), IM.INVOICE_DATE1))
        ELSE NULL
    END AS LastInvoiceFormatted,
    ROW_NUMBER() OVER (PARTITION BY IM.SYYY ORDER BY IM.INVOICE_DATE1 DESC) AS rn

    
FROM dbo.INVOICE_MASTER IM
JOIN ORDER_DISPATCH_INVOICE_LINK OL ON OL.IMASTERNO = IM.MASTERNO
JOIN ORDER_MASTER OM ON OL.OMASTERNO = OM.MASTERNO


)
SELECT
    LTRIM(RTRIM(s.SUP_NAME)) AS [Customer / Supplier],
    LTRIM(RTRIM(s.POSTCODE)) AS POSTCODE,
    LTRIM(RTRIM(s.CITY)) AS CITY,
    ISNULL(fin.TotalDebit, 0) AS [Total Debit],
    ISNULL(fin.TotalCredit, 0) AS [Total Credit],
    CASE 
        WHEN ISNULL(fin.TotalDebit, 0) = 0 THEN 0
        ELSE CEILING((ISNULL(fin.TotalCredit, 0) * 100.0) / ISNULL(fin.TotalDebit, 0))
    END AS [Payment %],
    (ISNULL(fin.TotalDebit, 0) - ISNULL(fin.TotalCredit, 0)) AS [Balance],
    ISNULL(CONVERT(varchar(10), p.LastPaymentDate, 103), 'No-Payment') AS [LastPaymentDate],
    CASE 
        WHEN p.LastPaymentDate IS NULL THEN 'No-Payment'
        ELSE 
            CASE 
                WHEN DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) < 0 THEN 'Error'
                WHEN DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) < 30 THEN 
                    CONVERT(varchar(10), DATEDIFF(DAY, p.LastPaymentDate, GETDATE())) + ' gün'
                ELSE 
                    CASE 
                        WHEN (DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) % 30) = 0 THEN
                            CONVERT(varchar(10), (DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) / 30)) + ' Ay'
                        ELSE
                            CONVERT(varchar(10), (DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) / 30)) + ' Ay ' +
                            CONVERT(varchar(10), (DATEDIFF(DAY, p.LastPaymentDate, GETDATE()) % 30)) + ' gün'
                    END
            END
    END AS [GecenZaman1],
    p.LastPaymentAmount,
    CASE
        WHEN p.LastPaymentType = 1 THEN '( CHEQUE )'
        WHEN p.LastPaymentType = 3 THEN '( Return )'
        WHEN p.LastPaymentType = 2 THEN 
            CASE
                WHEN p.CTYPE = 0 THEN 'Cash'
                WHEN p.CTYPE = 3 THEN 'TRANSFER'
                ELSE 'Bank/Cash'
            END
        WHEN p.LastPaymentType IS NULL THEN '( No Payment )'
        ELSE 'Other'
    END AS [Last Payment Type],
    ISNULL(CONVERT(varchar(10), li.LastInvoiceDate, 103), 'No-Invoice') AS [LastInvoiceDate],
    CASE 
        WHEN li.LastInvoiceDate IS NULL THEN 'No-Invoice'
        ELSE 
            CASE 
                WHEN DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) < 0 THEN 'Error'
                WHEN DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) < 30 THEN 
                    CONVERT(varchar(10), DATEDIFF(DAY, li.LastInvoiceDate, GETDATE())) + ' gün'
                ELSE 
                    CASE 
                        WHEN (DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) % 30) = 0 THEN
                            CONVERT(varchar(10), (DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) / 30)) + ' Ay'
                        ELSE
                            CONVERT(varchar(10), (DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) / 30)) + ' Ay ' +
                            CONVERT(varchar(10), (DATEDIFF(DAY, li.LastInvoiceDate, GETDATE()) % 30)) + ' gun'
                    END
            END
    END AS [GecenZaman2],
    LTRIM(RTRIM(ld.INVOICE_NO)) AS [Last Invoice No]

	,

	LTRIM(RTRIM(ld.ORDER_NO)) AS [Order No],
    CASE 
        WHEN LEFT(LTRIM(RTRIM(ld.ORDER_NO)), 3) = 'A01' THEN 'Hasan'
        WHEN LEFT(LTRIM(RTRIM(ld.ORDER_NO)), 3) = 'A02' THEN 'Murat'
        WHEN LEFT(LTRIM(RTRIM(ld.ORDER_NO)), 3) = 'A03' THEN 'Eren'
		WHEN LEFT(LTRIM(RTRIM(ld.ORDER_NO)), 3) = 'A04' THEN 'Serdar'
		WHEN LEFT(LTRIM(RTRIM(ld.ORDER_NO)), 3) = 'HA-' THEN 'HALIL'		
        ELSE 'OFIS'
    END AS [Sales Representative]


FROM dbo.SUPPLIER s
LEFT JOIN (
    SELECT 
        SYYY,
        SUM(DEBIT) AS TotalDebit,
        SUM(CREDIT) AS TotalCredit
    FROM dbo.ACC_TRANSACTION
    GROUP BY SYYY
) fin ON s.YYY = fin.SYYY
LEFT JOIN LastPaymentCTE p ON s.YYY = p.SYYY AND p.rn = 1
LEFT JOIN LastInvoiceCTE li ON s.YYY = li.SYYY
LEFT JOIN LastInvoiceDetailCTE ld ON s.YYY = ld.SYYY AND ld.rn = 1
WHERE s.ACCOUNT_TYPE = 0
AND s.SUP_NAME LIKE '%%'

-- 1 = 0 sadece balance'i olanlar
-- 0 = 0 tum musteriler balance olan olmayan! (eksi balance olanlar, sifir degerinden sonra basta gelir! )
-- ilk deger ? python icin paramdir!

AND (? = 0 OR (ISNULL(fin.TotalDebit, 0) - ISNULL(fin.TotalCredit, 0))  > 0)

ORDER BY 
    CASE 
        WHEN (ISNULL(fin.TotalDebit, 0) - ISNULL(fin.TotalCredit, 0)) > 0 THEN 1
        WHEN (ISNULL(fin.TotalDebit, 0) - ISNULL(fin.TotalCredit, 0)) < 0 THEN 2
        ELSE 3
    END,
    (ISNULL(fin.TotalDebit, 0) - ISNULL(fin.TotalCredit, 0)) DESC
