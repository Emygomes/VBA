create table ca200401 
(
	regiao_sigla varchar(2) NOT NULL,
	estado_sigla varchar(3) NOT NULL ,
	municipio varchar(15) NOT NULL ,
	Revenda varchar(50) NOT NULL ,
    CNPJ_da_Revenda	varchar(20)  NOT NULL,
	numero_rua int NOT NULL ,
	Bairro varchar(30) NOT NULL ,
	Cep varchar(18) NOT NULL ,
	Produto varchar(20) NOT NULL ,
	Data_da_coleta date,
	Valor_de_venda float(20) NOT NULL ,
	Valor_de_compra float(20) NOT NULL ,
	Unidade varchar(20) NOT NULL,
	Bandeira varchar(30) NOT NULL	
);

BULK INSERT ca200401
FROM 'C:\\ca200401.csv' /*alterar o local do arquivo */
WITH (
FIELDTERMINATOR = ';',
CODEPAGE = '65001',
ROWS_PER_BATCH = 1000
);

select*from ca200401

bulk insert ca200401 from 'C:\Temp\ca200401.csv' with (fieldterminator = ',', rowterminator = '\n', firstrow = 1, codepage = 'acp')

!!sqlcmd -S . -d curso -E -Q "set nocount on; select * from ca200401" -o "ca200401.csv" -W -s"," -h-1

drop table ca200401;

BULK INSERT ca200401 FROM 'C:\Temp\ca200401.csv' WITH (FIRSTROW=1,CODEPAGE='ACP',FIELDTERMINATOR = '',BATCHSIZE = 1000,ROWTERMINATOR = '\n')