CREATE TABLE dbo.tbl_TesteConnectionFactory(
	colString varchar(100),
	colNumeric numeric(10,2),
	colDate smalldatetime
)

--INSERT INTO tbl_TesteConnectionFactory (colString, colNumeric, colDate) VALUES ('Teste', 123.45, '2016-03-29')
--exec spc_TesteConnectionFactory 'Teste procedure',12345.67,'2016-03-29'

CREATE PROCEDURE dbo.spc_TesteConnectionFactory(
	@colString varchar(100),
	@colNumeric numeric(10,2),
	@colDate smalldatetime
)
AS
BEGIN

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED   
SET NOCOUNT ON  

INSERT INTO tbl_TesteConnectionFactory (colString, colNumeric, colDate) VALUES (@colString, @colNumeric, @colDate)

END


GRANT INSERT, UPDATE, DELETE ON tbl_TesteConnectionFactory TO DTI;
GRANT EXECUTE ON spc_TesteConnectionFactory TO DTI;
