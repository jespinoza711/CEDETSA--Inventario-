--USE DAS
--Create Database DASX

use CED2
/****** Object:  User [sysUser]    Script Date: 03/15/2014 22:33:08 ******/
--CREATE USER [sysUser] FOR LOGIN [sysUser] WITH DEFAULT_SCHEMA=[dbo]
--GO
/****** Object:  Table [dbo].[invTIPOMOVIMIENTO]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invTIPOMOVIMIENTO](
	[IDTipo] [int] NOT NULL,
	[Transaccion] [nvarchar](10) NULL,
	[Descr] [nvarchar](250) NULL,
	[Naturaleza] [nvarchar](1) NULL,
	[Factor] [smallint] NULL,
	[ReadOnly] [bit] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[invTIPOMOVIMIENTO] ADD CONSTRAINT PKTIPOMOVIMIENTO PRIMARY KEY ( IDTIPO )
GO

INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (1, 'COM', 'Compra', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (2, 'FAC', 'Facturación', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (3, 'AJE', 'Ajuste por Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (4, 'AJS', 'Ajuste por Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (5, 'BON', 'Bonificación', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (6, 'PRS', 'Préstamo Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (7, 'PRE', 'Préstamo Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (8, 'CON', 'Consumo', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (9, 'TRS', 'Traslado Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (10, 'TRE', 'Traslado Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (11, 'TRS', 'Traslado Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (12, 'FIE', 'Ajuste Físico Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (13, 'FIS', 'Ajuste Físico Salida', 'S', -1, 1)
GO

-- select * from [dbo].[invTIPOMOVIMIENTO]

/****** Object:  Table [dbo].[invPAQUETETIPOMOV]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invPAQUETE](
	[IDPaquete] [int] NOT NULL,
	PAQUETE NVARCHAR(20),
	[Descr] [nvarchar](250) NULL,
	[Consecutivo] [int] NULL,
	[ConsecAutomatico] [bit] NULL,
	[Documento] [nvarchar](20) NOT NULL,
	[Activo] [bit] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[invPAQUETE] ADD CONSTRAINT PKINVPAQUETE PRIMARY KEY (IDPAQUETE)
GO

INSERT [dbo].[invPAQUETE]  ( IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
SELECT IDTipo, Transaccion, 'Paquete de ' + Descr, 0, 1,  TRANSACCION + RIGHT( '00000000000'  + CAST( 0 AS NVARCHAR(20)),12), 1
from [dbo].[invTIPOMOVIMIENTO]


CREATE TABLE [dbo].[invPAQUETETIPOMOV](
	[IDPaquete] [int] NOT NULL,
	[IDTipo] [int] NOT NULL,
	[Transaccion] [nvarchar](10) NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[invPAQUETETIPOMOV] ADD CONSTRAINT PKPAQUETEMOV PRIMARY KEY (IDPAQUETE, IDTIPO)

ALTER TABLE [dbo].[invPAQUETETIPOMOV] ADD CONSTRAINT FKPAQUETEMOVPAQ FOREIGN KEY (IDPAQUETE) REFERENCES INVPAQUETE(IDPAQUETE)
GO
ALTER TABLE [dbo].[invPAQUETETIPOMOV] ADD CONSTRAINT FKPAQUETEMOVTIPOMOV FOREIGN KEY (IDTIPO) REFERENCES INVTIPOMOVIMIENTO(IDTIPO)
GO

INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
select P.IDPaquete, M.IDTipo, M.Transaccion 
from dbo.invPAQUETE P inner join dbo.invTIPOMOVIMIENTO M
on P.IDPaquete = M.IDTipo 
GO
-- SELECT * FROM [dbo].[invPAQUETETIPOMOV]
/****** Object:  Table [dbo].[invPAQUETE]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/****** Object:  Table [dbo].[invLOTE]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invLOTE](
	[IDLote] [int] NOT NULL,
	[LoteInterno] [nvarchar](50) NULL,
	[LoteProveedor] [nvarchar](50) NULL,
	[FechaVencimiento] [smalldatetime] NULL,
	[FechaFabricacion] [smalldatetime] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[invLOTE] ADD CONSTRAINT PKINVLOTE PRIMARY KEY (IDLOTE)
GO
/****** Object:  Table [dbo].[invBODEGA]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--DROP TABLE [dbo].[invBODEGA]
CREATE TABLE [dbo].[invBODEGA](
	[IDBodega] [int] NOT NULL,
	[Descr] [nvarchar](250) NULL,
	[Activo] [bit] NULL,
	[Factura] [bit] default 0,
	PreFactura nvarchar(20) default '', 
	ConsecFactura int default 0,
	ConsecPedido int default 0
 CONSTRAINT [PK_invBodega] PRIMARY KEY CLUSTERED 
(
	[IDBodega] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

Create Procedure dbo.invGetBodegas @IDBodega int
as
set nocount on
Select IDBodega, Descr DescrBodega, Activo, Factura, PreFactura , ConsecFactura , ConsecPedido 
From dbo.invBODEGA 
where (IDBodega = @IDBodega or @IDBodega = -1)
go
--use Ced2 drop procedure dbo.invUpdateBodega 
Create Procedure dbo.invUpdateBodega @Operacion nvarchar(1), @IDBodega int, @Descr nvarchar(255), @Activo bit, @Factura bit, @PreFactura nvarchar(20)
as
set nocount on
declare @NextCodigo int, @Ok bit
if @Operacion = 'I'
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IDBodega) + 1 from dbo.invBODEGA   ),1)
	while @Ok = 0
	begin
		if not exists (Select IDBodega from dbo.invBODEGA where IDBodega = @NextCodigo)
		begin

			insert dbo.invBODEGA  (IDBodega, Descr, Activo, Factura, PreFactura)
			values (@NextCodigo, @Descr, @Activo, @Factura, @PreFactura)
			set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IdBodega) + 1 from dbo.invBODEGA  ),1)		
		end
	end	
end
if @Operacion = 'U'
begin
 Update dbo.invBODEGA set Descr = @Descr, Activo = @Activo , Factura = @Factura
 where IDBodega = @IDBodega
end

if @Operacion = 'D'
begin
	Delete from dbo.invBODEGA 
	where IDBodega = @IDBodega
end
GO


/****** Object:  Table [dbo].[globalTABLAS]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE dbo.globalTABLAS(
	[IDTabla] [int] NOT NULL,
	[Nombre] [nvarchar](250) NULL,
	[Abrev] [nvarchar](20) NULL,
	[activo] [bit] NULL,
	[IDModulo] int not NULL default 0,
	[DescrUsuario] [nvarchar](250) NULL,
 CONSTRAINT [pkCatalogos] PRIMARY KEY CLUSTERED 
(
	[IDTabla] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

alter table dbo.globalTABLAS add constraint fkglobalTablas foreign key (IDModulo) references  secModulo (IDModulo)
go

/****** Object:  StoredProcedure [dbo].[spGlobalUpdateCatalogo]    Script Date: 03/15/2014 22:33:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Select * from [dbo].[globalCATALOGOS]


CREATE TABLE [dbo].[globalCATALOGOS](
	[IDCATALOGO] [nvarchar](20) NOT NULL,
	[IDTABLA] [int] NOT NULL,
	[CODIGO] [int] NOT NULL,
	[DESCR] [nvarchar](250) NULL,
	[ACTIVO] [bit] default 1,
	UsaValor bit default 0,
	NombreValor nvarchar(100) default 'ND',
	Valor decimal(8,2) default 0,
	Protected bit default 0,
	CodSistAnterior nvarchar(20)
	
 CONSTRAINT [PK2CATALOGOS] PRIMARY KEY CLUSTERED 
(
	[IDCATALOGO] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

alter table [globalCATALOGOS] add constraint fkglobalCATALOGOS foreign key (IDTAbla) references [dbo].[globalTABLAS] (idtabla)
go

----JULIO
--INSERT DBO.globalCATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
--VALUES ('ND', 0, 0, 'ND',1)
--GO

-- delete from [dbo].[globalCatalogos]  where IDTABLA <>0
--delete from [dbo].[globalTABLAS]  where IDTABLA <>0
/*
insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (0, 'Catalogo Global', 'GLOBAL',1,0, 'Admon')
go
*/

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (0, 'Catalogo Global', 'GLOBAL',1,0, 'Admon')
go
--DELETE from [dbo].[globalTABLAS] WHERE IDTABLA <> 0
insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (1, 'PRESENTACION', 'PRESENTACION',1,1000, 'Presentación')
go
insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (2, 'LINEA', 'CLASIF1',1,1000, 'Linea')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (3, 'FAMILIA', 'CLASIF2',1,1000, 'Familia')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (4, 'SUBFAMILIA', 'CLASIF3',1,1000, 'SubFamilia')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (5, 'PAIS', 'PAIS',1,0, 'País')
go
--select * from DBO.globalCATALOGOS
--INSERT DBO.globalCATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO, PROTECTED)
--VALUES ('ND', 5, 0, 'NICARAGUA',1, 1)
--GO

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (6, 'PLAZO', 'PLAZO',1,0, 'Plazo')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (7, 'MONEDA', 'MONEDA',1,0, 'Moneda')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (8, 'CATEGORIA_PROVEEDOR', 'CATPROVEEDOR',1,0, 'Categoría del Proveedor')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (9, 'IMPUESTO', 'IMPUESTO',1,0, 'Impuesto')
go


insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (10, 'TIPO_VENDEDOR', 'TIPO_VENDEDOR',1,0, 'Tipo Vendedor')
go
insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (11, 'TIPO_FACTURA', 'TIPO_FACTURA',1,0, 'Tipo Factura')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (12, 'DEPARTAMENTO', 'DEPATAMENTO',1,0, 'Departamento')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (13, 'MUNICIPIO', 'MUNICIPIO',1,0, 'Municipio')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (14, 'ZONA', 'ZONA',1,0, 'Zona')
go

insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (15, 'CATEGORIA_CLIENTE', 'CATEGORIA_CLIENTE',1,0, 'CategoriaCliente')
go
--insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
--values (15, 'CLASIFICACION_CLIENTE', 'CLASIF_CLIENTE',1,0, 'Clasificación Cliente')
--go SELECT * FROM  [dbo].[globalTABLAS] WHERE IDTABLA = 15
-- DELETE FROM dbo.globalCATALOGOS WHERE IDTABLA = 15

--drop trigger TRGCATALOGOS
CREATE TRIGGER TRGCATALOGOS 
   ON  DBO.globalCATALOGOS
   INSTEAD OF INSERT
AS 
BEGIN
	SET NOCOUNT ON;

	DECLARE @NEXT INT, @CODIGO INT, @IDTABLA INT, @DESCR NVARCHAR(250)
	
	SELECT @NEXT = MAX(ISNULL(T.CODIGO,0))+1
	FROM INSERTED I INNER JOIN globalCATALOGOS T 
	ON  I.IDTABLA = T.IDTABLA
	WHERE I.IDTABLA = T.IDTABLA
	IF @NEXT IS NULL
		SET @NEXT = 1
	
	INSERT DBO.globalCATALOGOS (IDCATALOGO, IDTABLA, CODIGO, DESCR, ACTIVO, UsaValor, NombreValor, Valor, Protected, CodSistAnterior)
	SELECT CAST(I.IDTABLA AS nvarchar(20)) +'-' + CAST (@NEXT AS NVARCHAR(20)),  I.IDTABLA, @NEXT, I.DESCR, I.ACTIVO, I.UsaValor, I.NombreValor, I.Valor, I.Protected, I.CodSistAnterior
	FROM INSERTED I 
END
GO

--select * from [cpproveedor]


/*
INSERT DBO.CATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
VALUES ('ND', 1, 0, 'CAJA-250-MG',1)
GO

INSERT DBO.CATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
VALUES ('ND', 1, 0, 'CANULA',1)
GO

INSERT DBO.CATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
VALUES ('ND', 2,0, 'UND',1)
GO

INSERT DBO.CATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
VALUES ('ND', 2,0, 'CAJA',1)
GO

INSERT DBO.CATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO)
VALUES ('ND',2,0, 'BULTO',1)
GO

drop procedure dbo.spGlobalUpdateCatalogo
*/

CREATE PROCEDURE [dbo].[spGlobalUpdateCatalogo] @Operacion nvarchar(1), @IDCatalogo nvarchar(20)='ND', @IDTable int, @Codigo int = 0, @Descr nvarchar(250), @Activo bit, @UsaValor bit, @NombreValor nvarchar(100), @Valor decimal(8,2)
as
set nocount on 
if UPPER(@Operacion) = 'I'
begin
	INSERT DBO.globalCATALOGOS(idcatalogo, IDTABLA, CODIGO,  DESCR, ACTIVO,UsaValor,NombreValor, Valor )
	VALUES (@IDCatalogo ,@IDTable,@Codigo, @Descr ,@Activo, @UsaValor, @NombreValor, @Valor )	
end
if UPPER(@Operacion) = 'U'
begin
	Update DBO.globalCATALOGOS set DESCR = @DESCR, ACTIVO = @Activo, UsaValor = @UsaValor , NombreValor = @NombreValor, Valor = @Valor 
	Where IDCATALOGO = @IDCatalogo
end

if UPPER(@Operacion) = 'D'
begin
	DELETE FROM DBO.globalCATALOGOS 
	Where IDCATALOGO = @IDCatalogo
end
GO


--drop procedure globalGetCatalogos
Create Procedure dbo.globalGetCatalogos @IDTabla int
as
set nocount on
Select C.IDCATALOGO, C.IDTABLA, T.nombre,  C.CODIGO, C.DESCR, C.Activo, C.UsaValor, C.NombreValor, C.Valor , C.protected
from DBO.globalCATALOGOS C inner join dbo.globalTABLAS T
on C.IDTABLA = T.IDTABLA 
Where (C.idtabla = @IDTabla or @IDTabla = -1 )
order by IDTabla, codigo
go

-- exec globalGetTablas -1 

Create Procedure globalGetTablas  @IDTabla int = -1 
as
set nocount on
Select T.IDTABLA, T.nombre, T.Abrev, T.IDModulo, M.Descr, T.DescrUsuario, T.activo
From dbo.globalTablas T inner join dbo.secmodulo M
on T.IDModulo = M.IDModulo
Where ( @IDTabla = -1 or T.IDTABLA = @IDTabla )
go

/****** Object:  View [dbo].[vGlobalCatalogo]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Table [dbo].[globalCATALOGOS]    Script Date: 03/15/2014 22:33:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Default [DF__invPAQUET__Conse__09DE7BCC]    Script Date: 03/15/2014 22:33:08 ******/
ALTER TABLE [dbo].[invPAQUETE] ADD  DEFAULT ((1)) FOR [ConsecAutomatico]
GO
-- select * from [dbo].[vGlobalCatalogo]
Create View [dbo].[vGlobalCatalogo]
as 
SELECT C.IDTABLA, T.nombre CATALOGO, T.DescrUsuario LABLE, T.IDMODULO, T.ABREV, C.CODIGO, C.IDCATALOGO,   C.DESCR, C.ACTIVO, C.UsaValor, C.NombreValor, C.Valor 
FROM globalCATALOGOS C INNER JOIN globalTABLAS T
ON C.IDTABLA = T.IDTABLA
GO
--drop table cpPROVEEDOR
Create Table cpPROVEEDOR( IDProveedor int not null, Nombre nvarchar(250), Telefono1 nvarchar(20), 
Telefono2 nvarchar(20), Telefono3 nvarchar(20), Fax nvarchar(20), Contacto nvarchar(250), TelefonoContacto nvarchar(20),
PorcDescuento decimal(8,2) default 0, PorcComisionAntes120  decimal(8,2) default 0, PorcComisionDespues120  decimal(8,2) default 0,
IDPais nvarchar(20), IDPlazo nvarchar(20), IDMoneda nvarchar(20),
IDCategoriaProveedor nvarchar(20), TrabajaconFOB BIT DEFAULT 0,
UserInsert nvarchar(20), UserUpdate nvarchar(20),
FechaInsert datetime, FechaUpdate datetime,
Activo bit default 1
)
GO


ALTER TABLE cpPROVEEDOR ADD CONSTRAINT pkcpProveedor primary key ( IDProveedor )
go 

Alter table  dbo.cpPROVEEDOR add constraint fkcpPROVEEDORPais foreign key (IDPais) references [globalCATALOGOS] ( IDCatalogo )
go

Alter table  dbo.cpPROVEEDOR add constraint fkcpPROVEEDORPlazo foreign key (IDPlazo) references [globalCATALOGOS] ( IDCatalogo )
go

Alter table  dbo.cpPROVEEDOR add constraint fkcpPROVEEDORMoneda foreign key (IDMoneda) references [globalCATALOGOS] ( IDCatalogo )
go

Alter table  dbo.cpPROVEEDOR add constraint fkcpPROVEEDORCategoria foreign key (IDCategoriaProveedor) references [globalCATALOGOS] ( IDCatalogo )
go


--drop table dbo.invPRODUCTO 
Create Table dbo.invPRODUCTO (	IDProducto int not null, Descr nvarchar(250), Impuesto nvarchar(20) ,
EsMuestra bit default 0, EsControlado bit default 0, Clasificacion1 nvarchar(20) , Clasificacion2 nvarchar(20) , 
Clasificacion3 nvarchar(20) ,EsEtico bit default 0, BajaPrecioDistribuidor bit default 0, 
IDProveedor int ,CostoUltLocal decimal(28,8) default 0, CostoUltDolar decimal(28,8) default 0, 
CostoUltPromLocal decimal(28,8) default 0, CostoUltPromDolar decimal(28,8) default 0,
PrecioPublicoLocal decimal(28,8) default 0, PrecioFarmaciaLocal decimal(28,8) default 0,
PrecioCIFLocal decimal(28,8) default 0, PrecioFOBLocal decimal(28,8) default 0,
IDPresentacion nvarchar(20), BajaPrecioProveedor bit default 0,
PorcDescAlzaProveedor decimal(8,2) default 0,
UserInsert nvarchar(20), UserUpdate nvarchar(20),
FechaInsert datetime, FechaUpdate datetime,
Activo bit default 1,
CodigoBarra nvarchar(128),
--BonifFAPorCada Decimal(28,8) default 0,
--'BonifFACantidad Decimal(28,8) default 0,
BonificaFA bit default 0,
BonifCOPorCada Decimal(28,8) default 0,
BonifCOCantidad Decimal(28,8) default 0 
)
go

Alter table  dbo.invPRODUCTO add constraint pkinvProducto primary key (IDProducto)
go

Alter table  dbo.invPRODUCTO add constraint fkImpuesto foreign key (Impuesto) references [globalCATALOGOS] ( IDCatalogo )
go
Alter table  dbo.invPRODUCTO add constraint fkClasificacion1 foreign key (Clasificacion1) references [globalCATALOGOS] ( IDCatalogo )
go
Alter table  dbo.invPRODUCTO add constraint fkClasificacion2 foreign key (Clasificacion2) references [globalCATALOGOS] ( IDCatalogo )
go
Alter table  dbo.invPRODUCTO add constraint fkClasificacion3 foreign key (Clasificacion3) references [globalCATALOGOS] ( IDCatalogo )
go
Alter table  dbo.invPRODUCTO add constraint fkPresentacion foreign key (IDPresentacion) references [globalCATALOGOS] ( IDCatalogo )
go

-- select * from dbo.vinvproducto
--Create View dbo.vinvProducto
--as
--Select IDProducto, Descr, Impuesto, case when EsMuestra = 1 then 'SI' ELSE 'NO' end EsMuestra,
--case when EsControlado = 1 then 'SI' ELSE 'NO' end EsControlado,
--case when EsEtico = 1 then 'SI' ELSE 'NO' end EsEtico,
--Clasificacion1, Clasificacion2, Clasificacion3,
--case when BajaPrecioDistribuidor = 1 then 'SI' ELSE 'NO' end BajaPrecioDistribuidor,
--case when BajaPrecioProveedor  = 1 then 'SI' ELSE 'NO' end BajaPrecioProveedor,
--IDProveedor, CostoUltLocal, CostoUltDolar, CostoUltPromLocal, CostoUltPromDolar, 
--PrecioPublicoLocal, PrecioFarmaciaLocal , PrecioCIFLocal, PrecioFOBLocal, 
--IDPresentacion, PorcDescAlzaProveedor, UserInsert, FechaInsert,  UserUpdate, FechaUpdate,
--CodigoBarra, case when Activo = 1 then 'SI' ELSE 'NO' end Activo
--From invProducto
--go
-- Drop procedure dbo.invUpdateProducto

CREATE PROCEDURE dbo.invUpdateProducto  @Operacion nvarchar(1),
	@IDProducto int  ,@Descr nvarchar(250) ,@Impuesto nvarchar(20) ,@EsMuestra bit ,@EsControlado bit ,@Clasificacion1 nvarchar(20) ,
	@Clasificacion2 nvarchar(20) ,@Clasificacion3 nvarchar(20) ,@EsEtico bit ,@BajaPrecioDistribuidor bit ,	@IDProveedor int ,
	@CostoUltLocal decimal(28, 8) ,	@CostoUltDolar decimal(28, 8) ,	@CostoUltPromLocal decimal(28, 8) ,	@CostoUltPromDolar decimal(28, 8) ,
	@PrecioPublicoLocal decimal(28, 8) ,	@PrecioFarmaciaLocal decimal(28, 8) ,	@PrecioCIFLocal decimal(28, 8) ,
	@PrecioFOBLocal decimal(28, 8) ,	@IDPresentacion nvarchar(20) ,	@BajaPrecioProveedor bit ,	@PorcDescAlzaProveedor decimal(8, 2) ,
	@UserInsert nvarchar(20) ,	@UserUpdate nvarchar(20) ,	@Activo bit ,	@CodigoBarra nvarchar(128), 
	@BonificaFA bit, @BonifCOPorCada decimal(28,8), @BonifCOCantidad decimal(28,8)
	
as
declare @NextCodigo int, @Ok bit
set nocount on
if @Operacion = 'I' -- se va a insertar 
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IdProducto) + 1 from dbo.invPRODUCTO  ),1)
	while @Ok = 0
	begin
		if not exists (Select IDProducto from dbo.invPRODUCTO where IDProducto = @NextCodigo)
		begin
			insert [dbo].[invPRODUCTO] ( [IDProducto] ,[Descr]  ,[Impuesto]  ,[EsMuestra] ,[EsControlado] ,[Clasificacion1]  ,[Clasificacion2]
				  ,[Clasificacion3] ,[EsEtico]  ,[BajaPrecioDistribuidor]  ,[IDProveedor]  ,[CostoUltLocal]  ,[CostoUltDolar]
				  ,[CostoUltPromLocal] ,[CostoUltPromDolar] ,[PrecioPublicoLocal] ,[PrecioFarmaciaLocal] ,[PrecioCIFLocal]
				  ,[PrecioFOBLocal] ,[IDPresentacion] ,[BajaPrecioProveedor] ,[PorcDescAlzaProveedor] ,[UserInsert]
				  ,[UserUpdate] ,[Activo] ,[CodigoBarra],
				  BonificaFA, BonifCOPorCada, BonifCOCantidad
			)
			VALUES ( @NextCodigo   ,@Descr ,@Impuesto  ,@EsMuestra  ,@EsControlado  ,@Clasificacion1  ,	@Clasificacion2 ,
				@Clasificacion3  ,	@EsEtico  ,	@BajaPrecioDistribuidor  ,	@IDProveedor  ,	@CostoUltLocal  ,	@CostoUltDolar  ,
				@CostoUltPromLocal  ,	@CostoUltPromDolar  ,	@PrecioPublicoLocal  ,	@PrecioFarmaciaLocal  ,	@PrecioCIFLocal  ,
				@PrecioFOBLocal  ,	@IDPresentacion  ,	@BajaPrecioProveedor  ,	@PorcDescAlzaProveedor  ,	@UserInsert  ,
				@UserUpdate  ,	@Activo  ,	@CodigoBarra,
				@BonificaFA, @BonifCOPorCada, @BonifCOCantidad
				)
		set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IdProducto) + 1 from dbo.invPRODUCTO  ),1)		
		end
	end
end
if @Operacion = 'U' -- se va a Actualizar
begin
	Update [dbo].[invPRODUCTO] set Descr = @Descr ,Impuesto = @Impuesto  ,EsMuestra =@EsMuestra  ,EsControlado = @EsControlado  
		,Clasificacion1 = @Clasificacion1  ,	Clasificacion2 = @Clasificacion2 ,Clasificacion3 = @Clasificacion3  ,
		EsEtico = @EsEtico  ,	BajaPrecioDistribuidor = @BajaPrecioDistribuidor  ,	IDProveedor = @IDProveedor  ,	
		CostoUltLocal = @CostoUltLocal  ,	CostoUltDolar = @CostoUltDolar  ,
		CostoUltPromLocal = @CostoUltPromLocal  , CostoUltPromDolar=@CostoUltPromDolar  ,PrecioPublicoLocal=	@PrecioPublicoLocal  
		,PrecioFarmaciaLocal = @PrecioFarmaciaLocal  ,	PrecioCIFLocal = @PrecioCIFLocal  ,
		PrecioFOBLocal = @PrecioFOBLocal  , IDPresentacion =	@IDPresentacion  ,BajaPrecioProveedor=	@BajaPrecioProveedor  ,	
		PorcDescAlzaProveedor = @PorcDescAlzaProveedor  , UserInsert =	@UserInsert  , UserUpdate=@UserUpdate  ,	
		Activo = @Activo  ,	CodigoBarra = @CodigoBarra ,
		BonificaFA = @BonificaFA, 
		BonifCOPorCada = @BonifCOPorCada, BonifCOCantidad = @BonifCOCantidad
	Where IDProducto = @IDProducto
end
if @Operacion = 'D' -- se va a Eliminar
begin
	DELETE FROM [dbo].[invPRODUCTO] Where IDProducto = @IDProducto
end

GO
-- vistas Catalogos
Create View dbo.vglobalCatalogos
as
SELECT C.IDCATALOGO, C.IDTABLA, T.nombre Tabla, C.CODIGO, C.DESCR, C.UsaValor,
C.NombreValor, C.Valor , T.IDModulo, c.ACTIVO 
FROM dbo.globalCATALOGOS C INNER JOIN dbo.globalTABLAS T
ON C.IDTABLA = T.IDTABLA 
go

-- VISTAS DE CATALOGOS
--DROP VIEW dbo.vinvImpuesto  SELECT * FROM dbo.vinvImpuesto 
Create View dbo.vinvImpuesto 
as
SELECT IDCATALOGO Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'IMPUESTO' and activo = 1
go

-- DROP VIEW dbo.vinvPresentacion 
Create View dbo.vinvPresentacion 
as
SELECT IDCATALOGO Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'PRESENTACION' and activo = 1
go

--DROP VIEW dbo.vinvClasificacion1 select * from dbo.vinvClasificacion1
Create View dbo.vinvClasificacion1
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'LINEA' and activo = 1
go

--DROP VIEW dbo.vinvClasificacion2
Create View dbo.vinvClasificacion2
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'FAMILIA' and activo = 1
go

--DROP VIEW dbo.vinvClasificacion3
Create View dbo.vinvClasificacion3
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'SUBFAMILIA' and activo = 1
go

--DROP VIEW dbo.vfafCategoriaVendedor
Create View dbo.vfafTipoVendedor
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'TIPO_VENDEDOR' and activo = 1
go

--DROP VIEW dbo.vinvClasificacion3
Create View dbo.vinvCategoriaProveedor
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'CATEGORIA_PROVEEDOR' and activo = 1
go

--DROP VIEW dbo.vinvClasificacion3 SELECT  Descr FROM dbo.vTipoFactura WHERE Codigo = '111'
--drop view dbo.vinvTipoFactura
Create View dbo.vfacTipoFactura
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'TIPO_FACTURA' and activo = 1
go


Create View dbo.vglobalImpuesto 
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo, G.UsaValor,G.NombreValor, G.Valor 
FROM dbo.vglobalCatalogos G
WHERE Tabla = 'IMPUESTO' and activo = 1
go

Create View dbo.vglobalMoneda
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo, G.UsaValor, G.NombreValor, G.Valor 
FROM dbo.vglobalCatalogos G
WHERE Tabla = 'MONEDA' and activo = 1
go

Create View dbo.vglobalPlazo
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo, G.UsaValor, G.NombreValor, G.Valor 
FROM dbo.vglobalCatalogos G
WHERE Tabla = 'PLAZO' and activo = 1
go

Create View dbo.vglobalCategoriaCliente
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo, G.UsaValor, G.NombreValor, G.Valor 
FROM dbo.vglobalCatalogos G
WHERE Tabla = 'CATEGORIA_CLIENTE' and activo = 1
go

Create View dbo.vglobalDepartamento
as
SELECT IDCATALOGO Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'DEPARTAMENTO' and activo = 1
go

Create View dbo.vglobalMunicipio
as
SELECT IDCATALOGO Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'MUNICIPIO' and activo = 1
go

Create View dbo.vglobalZona
as
SELECT IDCATALOGO Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'ZONA' and activo = 1
go
-- select * from vglobalCategoriaCliente drop view vglobalMoneda

-- drop view dbo.vinvProducto
--Create View dbo.vinvProducto
--as
--Select IDProducto, Descr, Impuesto, case when EsMuestra = 1 then 'SI' ELSE 'NO' end EsMuestra,
--case when EsControlado = 1 then 'SI' ELSE 'NO' end EsControlado,
--case when EsEtico = 1 then 'SI' ELSE 'NO' end EsEtico,
--Clasificacion1 Linea, Clasificacion2 Familia, Clasificacion3 SubFamilia,
--case when BajaPrecioDistribuidor = 1 then 'SI' ELSE 'NO' end BajaPrecioDistribuidor,
--case when BajaPrecioProveedor  = 1 then 'SI' ELSE 'NO' end BajaPrecioProveedor,
--IDProveedor, CostoUltLocal, CostoUltDolar, CostoUltPromLocal, CostoUltPromDolar, 
--PrecioPublicoLocal, PrecioFarmaciaLocal , PrecioCIFLocal, PrecioFOBLocal, 
--IDPresentacion, PorcDescAlzaProveedor, UserInsert, FechaInsert,  UserUpdate, FechaUpdate,
--CodigoBarra, case when Activo = 1 then 'SI' ELSE 'NO' end Activo
--From invProducto
--go


-- drop view dbo.vinvProducto
-- borrado el 04/05/2014
Create view dbo.vinvProducto
as
Select P.IDProducto 
      , P.Descr 
      , P.Impuesto,
		I.DESCR DescrImpuesto,
		I.Valor PorcImpuesto
      , case when P.EsMuestra = 1 then 'SI' ELSE 'NO' end EsMuestra
      , case when P.EsControlado = 1 then 'SI' ELSE 'NO' end EsControlado 
      , P.Clasificacion1 Linea, 
		L.descr DescrLinea
      , P.Clasificacion2 Familia,
		F.descr DescrFamilia		
      , P.Clasificacion3 SubFamilia,
		S.descr DescrSubFamilia
      , case when P.EsEtico = 1 then 'SI' ELSE 'NO' end EsEtico
      , case when P.BajaPrecioDistribuidor = 1 then 'SI' ELSE 'NO' end BajaPrecioDistribuidor 
      , P.IDProveedor,
		Pr.Nombre
      , P.CostoUltLocal 
      , P.CostoUltDolar 
      , P.CostoUltPromLocal 
      , P.CostoUltPromDolar 
      , P.PrecioPublicoLocal 
      , P.PrecioFarmaciaLocal 
      , P.PrecioCIFLocal 
      , P.PrecioFOBLocal 
      , P.IDPresentacion,
		Ps.descr DescrPresentacion
      , case when P.BajaPrecioProveedor  = 1 then 'SI' ELSE 'NO' end BajaPrecioProveedor 
      , P.PorcDescAlzaProveedor 
      , P.UserInsert 
      , P.UserUpdate 
      , P.FechaInsert 
      , P.FechaUpdate 
      , P.Activo 
      , P.CodigoBarra 
      , P.BonificaFA
      , P.BonifCOPorCada, P.BonifCOCantidad

From dbo.invPRODUCTO P inner join dbo.vinvClasificacion1 L 
on P.Clasificacion1 =  L.Codigo inner join dbo.vinvClasificacion2 F
on P.Clasificacion2 =  F.Codigo inner join dbo.vinvClasificacion3 S 
on P.Clasificacion3 =  S.Codigo inner join dbo.cpPROVEEDOR Pr
on P.IDProveedor = Pr.IDProveedor  inner join dbo.vinvPresentacion Ps
on P.IDPresentacion = Ps.Codigo INNER JOIN dbo.VGLOBALIMPUESTO I 
ON P.Impuesto = I.CODIGO
go 


--select * from dbo.vinvClasificacion1 exec dbo.invGetProductos -1
-- DROP PROCEDURE dbo.invGetProductos  select * from dbo.vinvProducto

Create Procedure dbo.invGetProductos @IDProducto int
as 
set nocount on 

Select P.IDProducto ,P.Descr  ,P.Impuesto  , I.Descr DescrImpuesto, P.EsMuestra ,P.EsControlado ,
                               P.Linea  , C1.Descr DescrClasif1, P.Familia, C2.Descr DescrClasif2
                               ,P.SubFamilia , C3.Descr DescrClasif3, P.EsEtico  ,P.BajaPrecioDistribuidor  ,P.IDProveedor , R.Nombre ,P.CostoUltLocal  ,P.CostoUltDolar
                                 , P.CostoUltPromLocal ,P.CostoUltPromDolar ,P.PrecioPublicoLocal ,P.PrecioFarmaciaLocal ,P.PrecioCIFLocal
                                 ,P.PrecioFOBLocal ,P.IDPresentacion , S.Descr DescrPresentacion, P.BajaPrecioProveedor ,P.PorcDescAlzaProveedor ,P.UserInsert
                                 ,P.UserUpdate ,P.Activo ,P.CodigoBarra, P.BonificaFA,
                                 P.BonifCOPorCada, P.BonifCOCantidad
                                
From dbo.vinvProducto P left join dbo.vinvImpuesto I 
on p.Impuesto = I.Codigo left join dbo.vinvClasificacion1 C1
on p.Linea=C1.Codigo left join dbo.vinvClasificacion2 C2
on p.Familia=C2.Codigo left join dbo.vinvClasificacion3 C3
on p.SubFamilia=C3.Codigo left join dbo.cpProveedor R
on p.IDProveedor = R.IDProveedor left join dbo.vinvPresentacion S
on P.IDPresentacion = S.codigo
where (@IDProducto = -1 or IDProducto = @IDProducto)

go 


--Create Procedure dbo.invGetProductos @IDProducto int
--as 
--set nocount on 

--Select P.IDProducto ,P.Descr  ,P.Impuesto  , I.Descr DescrImpuesto, P.EsMuestra ,P.EsControlado ,
--		P.Linea  , C1.Descr DescrClasif1, P.Familia , C2.Descr DescrClasif2
--		,P.SubFamilia , C3.Descr DescrClasif3, P.EsEtico  ,P.BajaPrecioDistribuidor  ,P.IDProveedor , R.Nombre ,P.CostoUltLocal  ,P.CostoUltDolar
--		  , P.CostoUltPromLocal ,P.CostoUltPromDolar ,P.PrecioPublicoLocal ,P.PrecioFarmaciaLocal ,P.PrecioCIFLocal
--		  ,P.PrecioFOBLocal ,P.IDPresentacion , S.Descr DescrPresentacion, P.BajaPrecioProveedor ,P.PorcDescAlzaProveedor ,P.UserInsert
--		  ,P.UserUpdate ,P.Activo ,P.CodigoBarra
--From dbo.vinvProducto P left join dbo.vinvImpuesto I 
--on p.Impuesto = I.Codigo left join dbo.vinvClasificacion1 C1
--on p.Linea=C1.Codigo left join dbo.vinvClasificacion2 C2
--on p.Familia=C2.Codigo left join dbo.vinvClasificacion3 C3
--on p.Subfamilia=C3.Codigo left join dbo.cpProveedor R
--on p.IDProveedor = R.IDProveedor left join dbo.vinvPresentacion S
--on P.IDPresentacion = S.codigo
--where (@IDProducto = -1 or IDProducto = @IDProducto)
--go

-- exec 


Create Table dbo.invEXISTENCIABODEGA ( IDBODEGA INT NOT NULL, IDPRODUCTO INT NOT NULL, 
	EXISTENCIA DECIMAL(28,4) DEFAULT 0, TRANSITO DECIMAL(28,4) DEFAULT 0 )
GO

ALTER TABLE dbo.invEXISTENCIABODEGA 
ADD CONSTRAINT pkEXistenciaBodega primary key (IDBODEGA, IDPRODUCTO)
GO

ALTER TABLE dbo.invEXISTENCIABODEGA ADD CONSTRAINT fkExistenciaBodega 
FOREIGN KEY (IDBODEGA ) REFERENCES invBODEGA (IDBODEGA)
GO

ALTER TABLE dbo.invEXISTENCIABODEGA ADD CONSTRAINT fkExistenciaBodegaPROD 
FOREIGN KEY (IDPRODUCTO ) REFERENCES invPRODUCTO (IDPRODUCTO)
GO
-- drop procedure dbo.invGetExistenciaBodega
--exec dbo.invGetExistenciaBodega -1, 1 exec dbo.invGetExistenciaBodega 1 , -1
Create Procedure dbo.invGetExistenciaBodega @IDProducto int, @IDBodega int 
as
set nocount on
Select E.IDBODEGA, B.Descr DescrBodega, E.IDPRODUCTO, P.Descr DescrProducto, E.EXISTENCIA, E.Transito
From dbo.invEXISTENCIABODEGA E inner join dbo.invBODEGA B
on E.IDBODEGA = B.IDBodega inner join dbo.invPRODUCTO P
on E.IDPRODUCTO = P.IDProducto 
Where ( E.IDPRODUCTO = @IDProducto or @IDProducto = -1)
and ( E.IDBODEGA  = @IDBodega or @IDBodega = -1)
go
-- DROP TABLE dbo.invEXISTENCIALOTE
Create Table dbo.invEXISTENCIALOTE ( IDBODEGA INT NOT NULL, IDPRODUCTO INT NOT NULL,
	IDLOTE INT NOT NULL, 
	EXISTENCIA DECIMAL(28,4) DEFAULT 0 )
GO

ALTER TABLE dbo.invEXISTENCIALOTE 
ADD CONSTRAINT pkEXistenciaLOTEBodega primary key (IDBODEGA, IDPRODUCTO, IDLOTE)
GO

ALTER TABLE dbo.invEXISTENCIALOTE ADD CONSTRAINT fkExistenciaBodegaLOTE 
FOREIGN KEY (IDBODEGA ) REFERENCES invBODEGA (IDBODEGA)
GO

ALTER TABLE dbo.invEXISTENCIALOTE ADD CONSTRAINT fkExistenciaBodegaPRODLOTE2 
FOREIGN KEY (IDPRODUCTO ) REFERENCES invPRODUCTO (IDPRODUCTO)
GO

ALTER TABLE dbo.invEXISTENCIALOTE ADD CONSTRAINT fkExistenciaBodegaPRODLOTE3
FOREIGN KEY (IDLOTE ) REFERENCES invLOTE (IDLOTE)
GO

-- drop table dbo.invCABMOVIMIENTOS
Create table dbo.invCABMOVIMIENTOS (
	IDPAQUETE INT NOT NULL,
	DOCUMENTO NVARCHAR(20) NOT NULL, 
	FECHA	 DATETIME NOT NULL, 
	CONCEPTO NVARCHAR(255),
	UserInsert nvarchar(20), UserUpdate nvarchar(20),
	FechaInsert datetime, FechaUpdate datetime	
)

GO
alter table dbo.invCABMOVIMIENTOS add constraint pkinvCabMovimientos primarY key (IDPAQUETE, DOCUMENTO, FECHA)
GO

-- DROP TABLE dbo.invMOVIMIENTOS
CREATE TABLE dbo.invMOVIMIENTOS ( IDPAQUETE INT NOT NULL, IDBODEGA INT NOT NULL, IDPRODUCTO INT NOT NULL, IDLOTE INT NOT NULL,
DOCUMENTO NVARCHAR(20) NOT NULL, FECHA DATETIME NOT NULL, IDTIPO INT NOT NULL, TRANSACCION NVARCHAR(10), NATURALEZA NVARCHAR(1),
CANTIDAD DECIMAL(28,4) DEFAULT 0, COSTOLOCAL DECIMAL(28,8) DEFAULT 0, 
COSTODOLAR DECIMAL(28,8) DEFAULT 0, PRECIOLOCAL DECIMAL(28,8) DEFAULT 0, 
PRECIODOLAR DECIMAL(28,8) DEFAULT 0, 
UserInsert nvarchar(20), UserUpdate nvarchar(20),
FechaInsert datetime, FechaUpdate datetime
)

go
alter table dbo.invMOVIMIENTOS add constraint pkMovimientos primary key (IDPAQUETE, IDBODEGA, IDPRODUCTO, IDLOTE, 
Documento, FECHA, IDTIPO)
GO

ALTER TABLE dbo.invMOVIMIENTOS  ADD CONSTRAINT fkMOVBODEGA 
FOREIGN KEY (IDBODEGA ) REFERENCES invBODEGA (IDBODEGA)
GO

ALTER TABLE dbo.invMOVIMIENTOS  ADD CONSTRAINT fkMOVPRODUCTO
FOREIGN KEY (IDPRODUCTO ) REFERENCES invPRODUCTO (IDPRODUCTO)
GO

ALTER TABLE dbo.invMOVIMIENTOS  ADD CONSTRAINT fkMOVLOTE
FOREIGN KEY (IDLOTE ) REFERENCES invLOTE (IDLOTE)
GO

ALTER TABLE dbo.invMOVIMIENTOS  ADD CONSTRAINT fkMOVTIPO
FOREIGN KEY (IDTIPO ) REFERENCES invTIPOMOVIMIENTO (IDTIPO)
GO

ALTER TABLE dbo.invMOVIMIENTOS  ADD CONSTRAINT fkMOVPAQ
FOREIGN KEY (IDPAQUETE ) REFERENCES invPAQUETE (IDPAQUETE)
GO

ALTER TABLE dbo.invMOVIMIENTOS ADD CONSTRAINT fkinvCABMOVIMIENTOS FOREIGN KEY (IDPAQUETE, Documento, Fecha ) references dbo.invCABMOVIMIENTOS (IDPAQUETE, Documento, Fecha )
go


-- para modulo Cuentas por Cobrar
-- drop table ccCLIENTE 
Create Table dbo.ccCLIENTE ( 
	IDCLIENTE int not null, Nombre nvarchar(255), RazonSocial nvarchar(255), Direccion nvarchar(255),
	Telefono1  varchar(20),	 Telefono2  varchar(20), Telefono3  varchar(20),
	Celular1  varchar(20),	 Celular2  varchar(20), email nvarchar(250),
	EsFarmacia bit default 0,
	NombreFarmacia nvarchar(250),
	RUC nvarchar(20),
	Propietario nvarchar(255),
	IDBodega int not null,
	IDPlazo nvarchar(20) not null,
	IDMoneda nvarchar(20) not null,
	IDCategoria nvarchar(20) not null,
	IDDepartamento nvarchar(20) not null,
	IDMunicipio nvarchar(20) not null,
	IDZona nvarchar(20) not null,
	IDVendedor int not null,
	FecUltCredito date,
	SALDOLOCAL decimal(28,8) default 0,
	SALDODOLAR decimal(28,8) default 0,
	TechoCredito decimal(28,8) default 0,  
	Activo bit default 1, 
	Credito bit default 1, 	
	UserInsert nvarchar(20), UserUpdate nvarchar(20),
	FechaInsert datetime, FechaUpdate datetime
	
	)
	
go
alter table ccCLIENTE ADD CONSTRAINT pkCliente primary key (IDCLIENTE )
go
--drop table ccDocumentos
Create Table ccDocumentos ( IDCLIENTE int not null,  Documento nvarchar(20) not null, Tipo nvarchar(3) not null, Fecha date not null, 
MontoLocal decimal(28,8) default 0, MontoDolar decimal(28,8) default 0,
Moneda nvarchar(10), FechaVencimiento date, TipoCambio decimal(28,8) default 0,
EsDebito bit default 0, Aprobado bit default 0
)
go
alter table ccDocumentos add constraint pkccDocumentos primary key clustered (IDCLIENTE, Documento, Tipo, Fecha)
go

--drop table ccDocumentosAplicacion
Create Table ccDocumentosAplicacion ( 
IDCLIENTE int not null,
DocDebito nvarchar(20) not null, TipoDebito nvarchar(3) not null,
DocCredito nvarchar(20) not null, TipoCredito nvarchar(3) not null,
FechaAplicacion Date not null,
MonedaDebito nvarchar(3) not null,
MonedaCredito nvarchar(3) not null,
MontoLocal decimal(28,8) default 0, 
MontoDolar decimal(28,8) default 0,
TipoCambio decimal(28,8) default 0,  Deslizamiento decimal(28,8) default 0,
SaldoInicDebitoLoc decimal(28,8) default 0, SaldoInicDebitoDol decimal(28,8) default 0,
SaldoFinalDebitoLoc decimal(28,8) default 0, SaldoFinalDebitoDol decimal(28,8) default 0,
SaldoInicCreditoLoc decimal(28,8) default 0, SaldoInicCreditoDol decimal(28,8) default 0,
SaldoFinalCreditoLoc decimal(28,8) default 0, SaldoFinalCreditoDol decimal(28,8) default 0

)
go

alter table ccDocumentosAplicacion add constraint pkccDocumentosAplicacion primary key clustered 
(IDCLIENTE, DocDebito, TipoDebito, DocCredito, TipoCredito)
go



-- Para gragar una Aplicación

--drop procedure dbo.ccUpdateAplicaciones
Create Procedure dbo.ccUpdateAplicaciones
@IDCLIENTE int, @DocDebito nvarchar(20), @TipoDebito nvarchar(3),
@DocCredito nvarchar(20), @TipoCredito nvarchar(3), 
@ValorCreditoDol decimal(28,8), @ValorCreditoLoc decimal(28,8), @Modalidad nvarchar(1),
@FechaCredito date, @TipoCambio decimal(28,8), @MonedaDebito nvarchar(3), @MonedaCredito nvarchar(3)
-- @Modalidad = 'I' Insertar , 'D' Delete 'R' Revertir Aplicacion
as
set nocount on
/*
set @IDCLIENTE =1
set @DocDebito = 'FAC1020'
set @TipoDebito = 'FAC'
SET @DocCredito = 'REC1214'
SET @TipoCredito = 'REC'
SET @ValorCreditoDol = 2000
SET @ValorCreditoLoc = 48606.00
SET @Modalidad = 'I'
SET @FechaCredito = '20140401'
set @TipoCambio = 24.303
exec dbo.ccUpdateAplicaciones 1, 'FAC1020', 'FAC', 'REC1215', 'REC', 1000, 24333, 'I', '20140405', 24.3333
*/

DECLARE @SaldoInicDebitoLoc decimal (28,8), @SaldoInicDebitoDol decimal (28,8),
@SaldoFinalDebitoLoc decimal (28,8), @SaldoFinalDebitoDol decimal (28,8),
@SaldoInicCreditoLoc decimal(28,8), @SaldoInicCreditoDol decimal(28,8),
@SaldoFinalCreditoLoc decimal(28,8), @SaldoFinalCreditoDol decimal(28,8)
-- PARA LOS CREDITOS
if not Exists (Select DocDebito -- SI NO EXISTE NINGUNA APLICACION A ESE DEBITO
			From ccDocumentosAplicacion 
			where IDCLIENTE = @IDCLIENTE and DocDebito = @DocDebito  and TipoDebito = @TipoDebito  ) 
BEGIN
	Select  top 1 @SaldoInicDebitoDol = MontoDolar, @SaldoFinalDebitoDol = MontoDolar - @ValorCreditoDol ,
	@SaldoInicDebitoLoc = MontoLocal, @SaldoFinalDebitoLoc = MontoLocal - @ValorCreditoLoc 
	From ccDocumentos
	Where IDCLIENTE = @IDCLIENTE and Documento = @DocDebito  and Tipo = @TipoDebito  
END
ELSE -- HAY AL MENOS UNA APLICACION A ESE DEBITO
BEGIN

	Select  top 1 @SaldoInicDebitoDol = SaldoFinalDebitoDol, @SaldoFinalDebitoDol = @SaldoInicDebitoDol - @ValorCreditoDol ,
	@SaldoInicDebitoLoc = SaldoFinalDebitoLoc, @SaldoFinalDebitoLoc = @SaldoInicDebitoLoc - @ValorCreditoLoc 
	From ccDocumentosAplicacion
	Where IDCLIENTE = @IDCLIENTE and DocDebito = @DocDebito  and TipoDebito = @TipoDebito  
	and FechaAplicacion <= @FechaCredito 
	order by IDCLIENTE, DocDebito, FechaAplicacion desc
	
END			

-- Actualizar el saldo del Doc tipo CREDITOS

if not Exists (Select DocCredito -- SI NO EXISTE NINGUNA APLICACION con ESE CREDITO
			From ccDocumentosAplicacion 
			where IDCLIENTE = @IDCLIENTE and DocCredito = @DocCredito  and TipoCredito = @TipoCredito  ) 
BEGIN
	Select Top 1 @SaldoInicCreditoDol = MontoDolar, @SaldoFinalCreditoDol = MontoDolar - @ValorCreditoDol ,
	@SaldoInicCreditoLoc = MontoLocal, @SaldoFinalCreditoLoc = MontoLocal - @ValorCreditoLoc 
	From ccDocumentos
	Where IDCLIENTE = @IDCLIENTE and Documento = @DocCredito  and Tipo = @TipoCredito  
END
ELSE -- HAY AL MENOS UNA APLICACION CON ESE CREDITO
BEGIN
	Select  top 1 @SaldoInicCreditoDol = SaldoFinalCreditoDol, @SaldoFinalCreditoDol = SaldoFinalCreditoDol - @ValorCreditoDol ,
	@SaldoInicCreditoLoc = SaldoFinalCreditoLoc, @SaldoFinalCreditoLoc = SaldoFinalCreditoLoc - @ValorCreditoLoc 
	From ccDocumentosAplicacion
	Where IDCLIENTE = @IDCLIENTE and DocCredito = @DocCredito  and TipoDebito = @TipoCredito
	and FechaAplicacion <= @FechaCredito 
	order by IDCLIENTE, DocCredito, FechaAplicacion desc
	
END			

if @Modalidad = 'I'
begin
	insert ccDocumentosAplicacion(IDCLIENTE, TipoDebito, DocDebito, TipoCredito, DocCredito,FechaAplicacion, 
	MontoDolar, MontoLocal,	
	SaldoInicDebitoDol, SaldoFinalDebitoDol, SaldoInicDebitoLoc, SaldoFinalDebitoLoc,
	SaldoInicCreditoDol, SaldoFinalCreditoDol, SaldoInicCreditoLoc,	SaldoFinalCreditoLoc,	
	TipoCambio, MonedaDebito, MonedaCredito )
	values (
		@IDCLIENTE, @TipoDebito, @DocDebito , @TipoCredito, @DocCredito,  @FechaCredito,
		@ValorCreditoDol, @ValorCreditoLoc, 
		@SaldoInicDebitoDol, @SaldoFinalDebitoDol,@SaldoInicDebitoLoc,@SaldoFinalDebitoLoc,  
		@SaldoInicCreditoDol, @SaldoFinalCreditoDol,@SaldoInicCreditoLoc,@SaldoFinalCreditoLoc,
		@TipoCambio, @MonedaDebito, @MonedaCredito )

end

go


Create Table dbo.fafVendedor ( IDVendedor int not null, Nombre nvarchar(255) not null, Activo bit default 0, 
Tipo nvarchar(20) not null)
go

alter table dbo.fafVendedor add constraint pkfafVendedor primary key (IDVendedor)
go
alter table dbo.fafVendedor add constraint fkVendedorTipo foreign key  (Tipo)  references dbo.globalCatalogos ( IDCatalogo )
go

--drop table dbo.fafPEDIDO alter table fafpedido add anulado bit default 0
-- drop table dbo.fafPEDIDO
CREATE TABLE dbo.fafPEDIDO
(
	IDPedido int not null,
	IDBodega int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	Aprobado bit default 0,
	BackOrder bit default 0,
	Anulado bit default 0,	
)
go
Alter Table dbo.fafPEDIDO add constraint pkfafPEDIDO primary key clustered ( IDPedido, IDBodega, IDCliente 
		 )
go		
Alter Table dbo.fafPedido add constraint fkfafPedidoCliente foreign key (IDCliente) references dbo.ccCliente (IDCLIENTE)
go

--Alter Table dbo.fafPedido add constraint fkfafPedidoVendedor foreign key (IDVendedor) references dbo.fafVendedor (IDVendedor)
--go

Alter Table dbo.fafPedido add constraint fkfafPedidoBodega foreign key (IDBodega) references dbo.invBodega (IDBodega)
go

Create Table dbo.fafPEDIDO_LINEA	-- drop table dbo.fafPEDIDO_LINEA
(
	IDPedido int not null,
	IDBodega int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	IDProducto int not null,
	IDLote int not null, 
	CantidadPedida decimal(28,8) default 0,
	CantidadFacturada decimal(28,8) default 0,
	CantidadNoFacturada decimal(28,8) default 0,
	Precio decimal(28,8) default 0,
	TotalImpuesto decimal(28,8) default 0,	
	TotalDescuento decimal(28,8) default 0,	
	SubTotal decimal(28,8) default 0,
	Total decimal(28,8) default 0,
	flgBonifica bit default 0,
	PorCada decimal(28,8) default 0,
	Bonifica decimal(28,8) default 0,
	flgDescuentoUnd bit default 0,
	flgDescuentoPorc bit default 0
)
go
Alter Table dbo.fafPEDIDO_LINEA add constraint pkfafPEDIDO_LINEA primary key clustered ( IDPedido, IDBodega, IDCliente, 
		IDProducto, IDLote )
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEA foreign key ( IDPedido, IDBodega, IDCliente 
		 ) references dbo.fafPedido ( IDPedido, IDBodega, IDCliente 
		)
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEABodega foreign key (IDBodega) references dbo.invBodega (IDBodega)
go


Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEAVendedor foreign key (IDVendedor) references dbo.fafVendedor (IDVendedor)
go


Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEACliente foreign key (IDCliente) references dbo.ccCliente (IDCLIENTE)
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEAProducto foreign key (IDProducto) references dbo.invProducto (IDProducto)
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEAProductoLote foreign key (IDLote) references dbo.invLote (IDLote)
go

-- drop table dbo.fafPEDIDO_LINEA 
-- drop table dbo.fafFACTURA
CREATE TABLE dbo.fafFACTURA
(
	IDBodega int not null,
	IDFactura int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	NoPreimpreso nvarchar(20),
	Anulada bit default 0,
	BackOrder bit default 0,
	IDPedido int default null,
	EsTeleventa bit default 0,
	TipoFactura nvarchar(20) not null
)
go
Alter Table dbo.fafFACTURA add constraint pkfafFACTURA primary key clustered ( IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha )
go

Alter Table dbo.fafFactura add Constraint fkfafFacturaCliente foreign key (IDCliente) references dbo.ccCliente( IDCLIENTE )
go

Alter Table dbo.fafFactura add Constraint fkfafFacturaBodega foreign key (IDBodega) references dbo.invBodega( IDBodega )
go

Alter Table dbo.fafFactura add Constraint fkfafFacturaVendedor foreign key (IDVendedor) references dbo.fafVendedor( IDVendedor )
go

Alter Table dbo.fafFactura add Constraint fkfafFacturaTipoFactura foreign key (TipoFactura) references dbo.globalCatalogos( IDCatalogo )
go

CREATE TABLE dbo.fafFACTURA_LINEA
(
	IDBodega int not null,
	IDFactura int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	IDProducto int not null,
	IDLOTE INT NOT NULL,
	Cantidad decimal(28,8) default 0,
	PrecioLocal decimal(28,8) default 0,
	PrecioDolar decimal(28,8) default 0,
	CostoLocal decimal(28,8) default 0,
	CostoDolar decimal(28,8) default 0,
	TipoCambio	decimal(28,8) default 0,
	SubTotalLocal decimal(28,8) default 0,
	SubTotalDolar decimal(28,8) default 0,
	SubImpuestoLocal decimal(28,8) default 0,
	SubImpuestoDolar decimal(28,8) default 0,
	TotalLocal decimal(28,8) default 0,
	TotalDolar decimal(28,8) default 0,
	FactorDevolucion smallint default 1, 
	CantidadDevuelta decimal(28,8) default 0
)
go

Alter Table dbo.fafFACTURA_LINEA add constraint pkfafFACTURA_LINEA primary key clustered ( IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha, IDProducto,IDLote )

GO

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEA foreign key ( IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha ) references dbo.fafFACTURA (IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha)
go		
Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEABodega foreign key (IDBodega)
references dbo.invBodega (IDBodega)
go

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEACliente foreign key (IDCliente)
references dbo.ccCLIENTE (IDCLIENTE)
go

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEAVendedor foreign key (IDVendedor)
references dbo.fafVendedor (IDVendedor)
go

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEAProducto foreign key (IDProducto)
references dbo.invProducto (IDProducto)
go

-- drop procedure dbo.fafUpdateVendedor 
Create Procedure dbo.fafUpdateVendedor @Operacion nvarchar(1), @IDVendedor int, @Nombre nvarchar(255), @Activo bit, @Tipo nvarchar(20)
as
set nocount on
declare @NextCodigo int, @Ok bit
if @Operacion = 'I'
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IDVendedor) + 1 from dbo.fafVendedor  ),1)
	while @Ok = 0
	begin
		if not exists (Select IDVendedor  from dbo.fafVendedor where IDVendedor = @NextCodigo)
		begin
 
			insert dbo.fafVendedor  (IDVendedor, Nombre, Activo, Tipo)
			values (@NextCodigo, @Nombre, @Activo, @Tipo)
			set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IDVendedor) + 1 from dbo.fafVendedor  ),1)		
		end
	end			
end
if @Operacion = 'U'
begin
 Update dbo.fafVendedor set Nombre = @Nombre, Activo = @Activo , Tipo= @Tipo
 where IDVendedor  = @IDVEndedor
end

if @Operacion = 'D'
begin
	Delete from dbo.fafVendedor
	where IDVendedor = @IDVendedor 
end
GO
-- JULIO

--insert dbo.fafVendedor ( IDVendedor , Nombre, tipo, Activo )
--values (1, 'Julio', '101', 1)
--drop procedure dbo.fafGetVendedores
Create Procedure dbo.fafGetVendedores @IDVendedor int
as
set nocount on
Select V.IDVendedor, V.Nombre , V.Tipo , C.Descr DescrTipo, V.Activo 
from dbo.fafVendedor V inner join dbo.vfafTipoVendedor C
on V.Tipo = C.Codigo
where (IDVendedor = @IDVendedor or @IDVendedor = -1)
go

-- drop procedure dbo.fafGetDetallePedidoToGrid
-- exec dbo.fafGetDetallePedidoToGrid -1, -1 exec dbo.fafGetDetallePedidoToGrid -1,-1
Create Procedure dbo.fafGetDetallePedidoToGrid @IDBodega int, @IDCliente int, @IDPedido int
as 
set nocount on

Create table #Resultado (IDPedido int, IDBodega int, IDCliente int, IDVendedor int, Fecha datetime, 
IDProducto int, IDLote int, Descr nvarchar(255), CantidadPedida decimal(28,8) default 0,
PrecioFarmaciaLocal decimal(28,8) default 0, IDImpuesto nvarchar(20),
PorcImpuesto decimal(28,8) default 0, Impuesto decimal(28,8) default 0,
 SubTotal decimal(28,8) default 0, TotalImpuesto decimal(28,8) default 0,
 Total decimal(28,8) default 0, Aprobado bit default 0, backOrder bit default 0, Anulado bit default 0
 )
 
insert #Resultado (IDPedido, IDBodega, IDCliente, IDVendedor, Fecha, IDProducto, IDLote, Descr, CantidadPedida, PrecioFarmaciaLocal, IDImpuesto,
 PorcImpuesto, Impuesto, SubTotal, TotalImpuesto, Total, Aprobado, BackOrder, Anulado )
Select P.IDPedido,  P.IDBodega, p.IDCliente,  P.IDVendedor,P.Fecha,  P.IDProducto, P.IDLote, A.Descr, P.CantidadPedida, A.PrecioFarmaciaLocal,
 A.Impuesto IDImpuesto, A.PorcImpuesto,  A.PrecioFarmaciaLocal* (A.PorcImpuesto/100) Impuesto, 
 (P.CantidadPedida* A.PrecioFarmaciaLocal) SubTotal ,
((P.CantidadPedida* A.PrecioFarmaciaLocal)* A.PorcImpuesto/100) TotalImpuesto,

( (P.CantidadPedida* A.PrecioFarmaciaLocal) +
((P.CantidadPedida* A.PrecioFarmaciaLocal)* A.PorcImpuesto/100)) Total,
D.Aprobado, D.BackOrder, D.Anulado

From dbo.fafPEDIDO_Linea P inner join dbo.vinvProducto A
on P.IDProducto = A.IDProducto inner join dbo.fafPedido D
on P.IDPedido = D.IDPedido and P.idBodega = D.IDBodega and P.IDCliente = D.IDCliente
and P.IDVendedor = D.IDVendedor
where P.IDBODEGA = @IDBodega AND P.IDCliente = @IDCliente and P.IDPedido = @IDPedido 

Select R.IDPedido, R.IDBodega, B.DESCR DESCRBODEGA, R.IDCliente, C.NOMBRE, R.IDVendedor, V.Nombre DescrVendedor,
R.Fecha, R.IDProducto, R.Descr, R.IDLote, R.CantidadPedida, R.PrecioFarmaciaLocal, R.IDImpuesto,
 R.PorcImpuesto, R.Impuesto, R.SubTotal, R.TotalImpuesto, R.Total, 
 R.Aprobado, R.BackOrder, R.Anulado
 From #Resultado R INNER JOIN DBO.ccCLIENTE C ON R.IDCLIENTE = C.IDCLIENTE 
 INNER JOIN DBO.INVBODEGA B ON R.IDBODEGA = B.IDBODEGA 
 INNER JOIN DBO.FAFVENDEDOR V ON R.IDVENDEDOR = V.IDVENDEDOR
 
DROP TABLE #RESULTADO
GO
---*********************************** drop procedure dbo.fafUpdatePedido

Create Procedure dbo.fafUpdatePedido @Operacion nvarchar(1), @IDPedido int , @IDBodega int, @IDCliente int, @IDVendedor int, @Fecha datetime, 
@Aprobado bit, @BackOrder bit, @Anulado bit
as
set nocount on
declare @NextCodigo int, @Ok bit
if @Operacion = 'I'
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(ConsecPedido) + 1 from dbo.invBODEGA (UPDLOCK) where IDBodega= @IDBodega  ),1)
	while @Ok = 0
	begin
		if not exists (Select ConsecPedido  from dbo.invBODEGA where IDBodega= @IDBodega and  ConsecPedido = @NextCodigo)
		begin
			insert dbo.fafPEDIDO ( IDPedido, IDBodega , IDCliente, IDVendedor, Fecha, Aprobado, BackOrder, Anulado )
			values ( @NextCodigo,@IDBodega, @IDCliente, @IDVendedor, @Fecha, @Aprobado, @BackOrder, @Anulado )
			set @Ok = 1
			update dbo.invBODEGA set ConsecPedido = @NextCodigo Where IDBodega= @IDBodega
			SET @IDPedido = @NextCodigo
			
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(ConsecPedido) + 1 from dbo.invBODEGA (UPDLOCK) where IDBodega= @IDBodega   ),1)		
		end
	end			
end
if @Operacion = 'D'
begin
	Delete from dbo.fafPEDIDO_LINEA where IDPedido = @IDPedido and IDCliente = @IDCliente and IDBodega = @IDBodega 
	Delete from dbo.fafPEDIDO where IDPedido = @IDPedido and IDCliente = @IDCliente and IDBodega = @IDBodega 
end

if @Operacion = 'U'
begin
	Update dbo.fafPEDIDO set Anulado = @Anulado, Aprobado = @Aprobado, BackOrder = @BackOrder 
	where IDPedido = @IDPedido and IDCliente = @IDCliente and IDBodega = @IDBodega 
end

SELECT @IDPedido IDPEDIDO
go
-- DROP PROCEDURE dbo.fafUpdatePedidoLinea -- Select * from  dbo.fafPEDIDO_LINEA
Create Procedure dbo.fafUpdatePedidoLinea @Operacion nvarchar(1), @IDPedido int, @IDBodega int, @IDCliente int, @IDVendedor int, @Fecha datetime, 
@IDProducto int, @IDLote int, @CantidadPedida decimal(28,8), @Precio decimal(28,8), @SubTotal decimal(28,8), @TotalDescuento decimal(28,8),  
@TotalImpuesto decimal(28,8), @Total decimal(28,8),
@flgBonifica bit, @PorCada int, @Bonifica bit, @flgDescuentoUnd bit,  @flgDescuentoPorc bit
as
set nocount on
declare @NextCodigo int, @Ok bit
if @Operacion = 'I'
begin
	insert dbo.fafPEDIDO_LINEA ( IDPedido,IDBodega , IDCliente, IDVendedor, Fecha, IDProducto, IDLote, CantidadPedida, Precio, SubTotal, TotalDescuento,  TotalImpuesto, Total, flgBonifica , PorCada , Bonifica , flgDescuentoUnd ,  flgDescuentoPorc  )
	values ( @IDPedido, @idbodega, @IDCliente, @IDVendedor, @Fecha, @IDProducto, @IDLote,  @CantidadPedida, @Precio, @SubTotal, @TotalDescuento, @TotalImpuesto, @Total,@flgBonifica , @PorCada , @Bonifica , @flgDescuentoUnd ,  @flgDescuentoPorc  )
end
if @Operacion = 'D'
begin
	Delete from dbo.fafPEDIDO_LINEA  where IDPedido = @IDPedido 
end

go

-- DROP TABLE dbo.invbodegaUsuario 
Create Table dbo.invbodegaUsuario ( IDBodega int NOT NULL, Usuario nvarchar(20) NOT NULL, Factura bit default 0, ConsultaInv bit default 0 )
go

alter table dbo.invbodegaUsuario add constraint pkinvbodegaUsuario primary key clustered ( IDBodega, Usuario )
go

alter table dbo.invbodegaUsuario add constraint fkinvbodegaUsuarioBodega foreign key (IDBodega)  references dbo.invBodega (IDBodega)
go

alter table dbo.invbodegaUsuario add constraint fkinvbodegaUsuariouSER foreign key (Usuario)  references dbo.secUsuario (Usuario)
go


CREATE PROCEDURE dbo.invUpdateBodegaUsuario  @Operacion nvarchar(1), @IDBodega int , @Usuario nvarchar(20) , @Factura bit , @ConsultaInv bit 
AS
Set Nocount on 

if @Operacion = 'I'
begin
	INSERT dbo.invbodegaUsuario ( IDBodega , Usuario , Factura , ConsultaInv  )
	VALUES ( @IDBodega , @Usuario , @Factura , @ConsultaInv  )
end
if @Operacion = 'U'
begin
	update dbo.invbodegaUsuario set  Factura = @Factura, ConsultaInv = @ConsultaInv 
	where IDbodega = @IDBodega  and Usuario = @Usuario
end
if @Operacion = 'D'
begin
	DELETE FROM dbo.invbodegaUsuario 
	where IDbodega = @IDBodega  and Usuario = @Usuario
end

GO

Create view dbo.vinvBodegaUsuario
as 
Select bu.IDBodega, b.Descr DescrBodega, bu.Usuario, u.DESCR Nombre, bu.Factura, bu.ConsultaInv 
From dbo.invbodegaUsuario bu inner join dbo.invBODEGA b
on bu.IDBodega = b.IDBodega inner join dbo.secUSUARIO u
on bu.Usuario = u.USUARIO 
go
-- exec dbo.getBodegaUsuario 1 drop procedure dbo.getBodegaUsuario
CREATE PROCEDURE dbo.getBodegaUsuario @IDBodega int
as
set nocount on
Select IDBodega, DescrBodega, Usuario, Nombre, Factura, ConsultaInv 
From  dbo.vinvBodegaUsuario
Where (IDBodega = @IDBodega or  @IDBodega  = -1) 
go
-- drop function dbo.fafgetCantBodegaFacturableForUser 
CREATE FUNCTION dbo.fafgetCantBodegaFacturableForUser ( @Usuario nvarchar(20) ) 
returns int
begin

declare @Resultado int
set @Resultado = (
Select COUNT(*) 
From  dbo.vinvBodegaUsuario
Where (Usuario = @usuario) and Factura =1
)
if @Resultado is null
	set @Resultado = 0
return @Resultado
end
go
-- select * from dbo.vfafPedido drop view dbo.vfafPedido 
-- drop view dbo.vfafPedido 
Create View dbo.vfafPedido 
as 
Select P.IDPedido, P.Fecha, P.IDBodega, B.Descr DescrBodega,
P.IDCliente, C.Nombre , P.IDVendedor, V.Nombre DescrVendedor, L.IDProducto, A.Descr DescrProducto,
L.IDLote, T.LoteProveedor, L.CantidadPedida, L.CantidadFacturada, L.CantidadNoFacturada, L.Precio, 
(L.CantidadPedida * L.Precio ) SubTotal, 
(L.CantidadPedida * L.Precio ) *  A.PorcImpuesto/100 TotalImpuesto, 
(L.CantidadPedida * L.Precio ) + ((L.CantidadPedida * L.Precio ) *  A.PorcImpuesto/100) Total,
P.Anulado, P.BackOrder , P.Aprobado 
from dbo.fafPEDIDO P inner join dbo.fafPEDIDO_LINEA L
on P.IDPedido = L.IDPedido and P.IDCliente = L.IDCliente and P.IDBodega =L.IDBodega 
and P.IDVendedor = L.IDVendedor
inner join dbo.invBODEGA B on  P.IDBodega = B.IDBodega 
Inner join dbo.ccCLIENTE C on   P.IDCliente = C.IDCLIENTE
inner join dbo.fafVendedor V on P.IDVendedor = V.IDVendedor
inner join dbo.vinvPRODUCTO A on L.IDProducto = A.IDProducto 
inner join dbo.invLOTE T on L.IDLote = T.IDLote 

go

-- exec fafgetPedidos 'C', 5,10, '20140101', '20500101', 1, -1, 1,0
--drop procedure dbo.fafgetPedidos
Create Procedure  dbo.fafgetPedidos @Modo nvarchar(1), @Pedido1 int , @Pedido2 int,  @FechaInic datetime, @FechaFin datetime,
@IDCliente int,  @IDVendedor int, @Desaprobados smallint, @Anuladas smallint
as
set nocount on
if @Modo = 'D' -- Detallado
begin
	Select  [IDPedido]  ,[Fecha]  ,[IDBodega]  ,[DescrBodega]  ,[IDCliente]  ,[Nombre]
		  ,[IDVendedor]  ,[DescrVendedor], IDProducto,   DescrProducto, IDLote, LoteProveedor, [CantidadPedida]  ,[CantidadFacturada] ,[CantidadNoFacturada]
		  ,[Precio] ,[SubTotal] ,[TotalImpuesto] ,[Total] ,[Anulado] ,[BackOrder]   ,[Aprobado]
	From dbo.vfafPedido 
	where ( IDPedido between @Pedido1 and @Pedido2 or (@Pedido1+@Pedido2)=0)
	and ( Fecha between @FechaInic and @FechaFin )
	and ( IDCliente = @IDCliente or @IDCliente = -1)
	and ( IDVendedor = @IDVendedor or @IDVendedor = -1)
	and ( (Aprobado = 0 and @Desaprobados = 1 )or (@Desaprobados=0 and Aprobado in (0,1)))
	and ( (Anulado = 0 and @Anuladas = 1 )or (@Anuladas=0 and Anulado in (0,1)))
	order by Fecha, IDCliente, IDPedido, IDProducto
end
if @Modo = 'C' -- Detallado
begin
	Select  [IDPedido]  ,[Fecha]  ,[IDBodega]  ,[DescrBodega]  ,[IDCliente]  ,[Nombre]
		  ,[IDVendedor]  ,[DescrVendedor],  Anulado ,[Aprobado],
		  SUM( SUBTOTAL) SUBTOTAL ,
		  SUM([TotalImpuesto]) TotalImpuesto ,Sum([Total]) Total  
	From dbo.vfafPedido 
	where ( IDPedido between @Pedido1 and @Pedido2 or (@Pedido1+@Pedido2)=0)
	and ( Fecha between @FechaInic and @FechaFin )
	and ( IDCliente = @IDCliente or @IDCliente = -1)
	and ( IDVendedor = @IDVendedor or @IDVendedor = -1)
	and ( (Aprobado = 0 and @Desaprobados = 1 )or (@Desaprobados=0 and Aprobado in (0,1)))
	and ( (Anulado = 1 and @Anuladas = 1 )or (@Anuladas=0 and Anulado in (0,1)))
	group by   [IDPedido]  ,[Fecha]  ,[IDBodega]  ,[DescrBodega]  ,[IDCliente]  ,[Nombre]
		  ,[IDVendedor]  ,[DescrVendedor], [Anulado] ,[Aprobado]
	order by  IDCliente, fecha,  IDPedido
end

go
-- BEGIN JULIO
-- AGREGADO POR JULIO 03/05/2014
delete from dbo.[invPAQUETETIPOMOV]
go 
delete from dbo.invTIPOMOVIMIENTO
go 
delete from dbo.invPAQUETE
go
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (1, 'COM', 'Compra', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (2, 'FAC', 'Facturación', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (3, 'AJE', 'Ajuste por Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (4, 'AJS', 'Ajuste por Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (5, 'BON', 'Bonificación', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (6, 'PRS', 'Préstamo Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (7, 'PRE', 'Préstamo Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (8, 'CON', 'Consumo', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (9, 'TRS', 'Traslado Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (10, 'TRE', 'Traslado Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (11, 'FIE', 'Ajuste Físico Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (12, 'FIS', 'Ajuste Físico Salida', 'S', -1, 1)
GO


--INSERTAR PAQUETES
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (1,'COM','Paquete de Compra',0,1,'COM000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (2,'FAC','Paquete de Facturación',0,1,'FAC000000000000',1)
GO
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (3,'AJU','Paquete de Ajuste',0,1,'AJU000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (4,'BON','Paquete de Bonificación',0,1,'BON000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (5,'PRE','Paquete de Préstamo',0,1,'PRE000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (6,'CON','Paquete de Consumo',0,1,'CON000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (7,'TRS','Paquete de Traslado',0,1,'TRS000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (8,'FIS','PPaquete de Ajuste Físico',0,1,'FIS000000000000',1)
GO

--INSERTAR PAQUETE - TIPO MOVIMIENTO
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (1,1,'COM')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (2,2,'FAC')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (3,3,'AJU')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (3,4,'AJU')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (4,5,'BON')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (5,6,'PRS')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (5,7,'PRS')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (6,8,'CON')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (7,9,'TRS')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (7,10,'TRS')
GO 

--SELECT * FROM dbo.invTIPOMOVIMIENTO
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (8,11,'FIS')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (8,12,'FIS')
GO


ALTER TABLE dbo.invCABMOVIMIENTOS ADD REFERENCIA NVARCHAR (20) NOT NULL

GO 




-- Drop procedure dbo.invInsertMovimientos
CREATE PROCEDURE dbo.invInsertMovimientos  @IDPaquete int,
	@IDBodega int  ,@IDProducto int,@IDLote INT ,@Documento nvarchar(20),@Fecha DATETIME,@IdTipo INT,
	@Transaccion NVARCHAR(10), @Naturaleza AS NVARCHAR(1),@Cantidad DECIMAL(28,8) ,@CostoDolar DECIMAL(28,8),
	@CostoLocal DECIMAL(28,8),@PrecioDolar DECIMAL(28,8), @PrecioLocal  DECIMAL(28,8),@UserInsert AS NVARCHAR(20),
	@UserUpdate  NVARCHAR(20)   
	
as

set nocount on


	insert [dbo].invMOVIMIENTOS( IdPaquete,IDBodega,IDProducto,IDLote,Documento,Fecha,IDTipo,Transaccion,Naturaleza,
			Cantidad,CostoDolar,CostoLocal,PrecioDolar,PrecioLocal,UserInsert,UserUpdate,FechaInsert,FechaUpdate)
	
	VALUES ( @IDPaquete,@IDBodega,@IDProducto,@IDLote  ,@Documento  ,@Fecha  ,	@IdTipo ,
		@Transaccion  ,	@Naturaleza,@Cantidad,@CostoDolar,@CostoLocal,@PrecioDolar,@PrecioLocal,
		@UserInsert,@UserUpdate,GETDATE(),GETDATE()	)



GO
--DROP PROCEDURE dbo.invGetNextConsecutivoPaquete
CREATE PROCEDURE dbo.invGetNextConsecutivoPaquete(@IDPaquete AS INT ,@Documento AS NVARCHAR(20) OUTPUT)
AS 

	DECLARE @NextCodigo AS INT
	DECLARE @NextCodigoString AS NVARCHAR(20)

	SELECT @NextCodigo=isnull(Consecutivo,0)+ 1,@NextCodigoString=  Paquete + RIGHT('000000000000'  + Cast(isnull(Consecutivo,0) + 1 AS NVARCHAR(20)),12) 
	FROM DBO.invPAQUETE  WHERE IDPaquete=@IDPaquete
	
	UPDATE dbo.invPAQUETE SET Consecutivo = @NextCodigo,Documento = @NextCodigoString WHERE IDPaquete=@IDPaquete
	
	set @Documento =@NextCodigoString
	
	
GO 

--DROP PROCEDURE dbo.invInsertCabMovimientos
CREATE  PROCEDURE [dbo].[invInsertCabMovimientos]  @IDPaquete int,
	@Documento nvarchar(20) OUTPUT,@Fecha DATETIME,@Concepto NVARCHAR(255),@Referencia  nvarchar(20),@UserInsert AS NVARCHAR(20),
	@UserUpdate  NVARCHAR(20),@ActualizaConsecutivo AS BIT
	
as

set nocount ON
IF (@ActualizaConsecutivo=1)  
	exec dbo.invGetNextConsecutivoPaquete @IDPaquete,@Documento OUTPUT

INSERT INTO dbo.invCABMOVIMIENTOS(IDPAQUETE, DOCUMENTO, FECHA, CONCEPTO,
            UserInsert, UserUpdate, FechaInsert, FechaUpdate,REFERENCIA)

VALUES (@IDPaquete,@Documento,@Fecha,@Concepto,@UserInsert,@UserUpdate,GETDATE(),GETDATE(),@Referencia )


SELECT @Documento Documento
GO 

--DROP PROCEDURE dbo.invGetCabeceraDocumento 
CREATE PROCEDURE dbo.invGetCabeceraDocumento @IDPaquete AS INT, @Documento AS NVARCHAR(20),@FechaInicial AS DATETIME,@FechaFinal AS DATETIME
AS 

set nocount ON

SELECT cm.IDPAQUETE,p.PAQUETE,p.Descr DescrPaquete ,cm.DOCUMENTO, cm.FECHA, cm.CONCEPTO, cm.UserInsert,
       cm.UserUpdate, cm.FechaInsert, cm.FechaUpdate
FROM dbo.invCabMovimientos cm
INNER JOIN dbo.invPaquete p ON cm.IDPAQUETE=p.IDPaquete
WHERE (cm.IdPaquete = @IDPaquete OR @IDPaquete=-1)
AND (cm.Documento=@Documento OR @Documento='*')
AND cm.FechaInsert between @FechaInicial AND @FechaFinal

 
GO


CREATE VIEW vinvMovimientos
AS 
SELECT	mov.IDPAQUETE, 
		p.PAQUETE,
		p.Descr DescrPaquete,
		mov.IDBODEGA, 
		b.Descr DescrBodega,
		mov.IDPRODUCTO,
		vp.Descr DescrProducto,
		mov.IDLOTE, 
		L.LoteInterno, 
		L.LoteProveedor, 
		L.FechaVencimiento, 
		L.FechaFabricacion,
		mov.DOCUMENTO,
		mov.FECHA, 
		mov.IDTIPO, 
		TM.Descr DescrTipo,
		TM.Factor,
		mov.TRANSACCION, 
		mov.NATURALEZA, 
		mov.CANTIDAD,
		mov.COSTOLOCAL, 
		mov.COSTODOLAR, 
		mov.PRECIOLOCAL, 
		mov.PRECIODOLAR,
		mov.UserInsert, 
		mov.UserUpdate, 
		mov.FechaInsert, 
		mov.FechaUpdate
FROM dbo.invMOVIMIENTOS mov INNER JOIN dbo.invPAQUETE P ON mov.IDPAQUETE=p.IDPaquete
INNER JOIN dbo.invBODEGA B ON mov.IDBODEGA=B.IDBodega
INNER JOIN dbo.vinvProducto vp ON mov.IDPRODUCTO=vp.IDProducto
INNER JOIN dbo.invLOTE L ON mov.IDLOTE=l.IDLote
INNER JOIN dbo.invTIPOMOVIMIENTO TM ON mov.IDTIPO=TM.IDTipo

GO 


CREATE PROCEDURE [dbo].[invGetDetalleMovimiento](@IDPaquete AS INT,@Documento AS NVARCHAR(20))
AS 


SELECT IDBODEGA BodegaOrigen,DescrBodega DescrBodegaOrigen,IDBodega BodegaDestino,DescrBodega DescrBodegaDestino, IDPRODUCTO, DescrProducto, IDLOTE, LoteInterno,
       FechaVencimiento, FechaFabricacion, DOCUMENTO, FECHA, IDTIPO, DescrTipo,
       TRANSACCION, NATURALEZA, CANTIDAD, COSTOLOCAL, COSTODOLAR, PRECIOLOCAL,
       PRECIODOLAR, UserInsert
  FROM dbo.vinvMovimientos
WHERE IDPAQUETE=@IDPaquete AND DOCUMENTO=@Documento 

SELECT * FROM dbo.vinvMovimientos

GO 


CREATE VIEW dbo.vinvPaqueteTipoMovimiento
AS 
SELECT p.IDPaquete, p.PAQUETE, p.Descr DescrPaquete,tm.IDTipo, tm.Transaccion, tm.Descr DescrTipo,
       tm.Naturaleza, tm.Factor, tm.[ReadOnly] 
FROM dbo.invPAQUETE p
INNER JOIN dbo.invPAQUETETIPOMOV pm ON pm.IDPaquete = p.IDPaquete
INNER JOIN dbo.invTIPOMOVIMIENTO tm ON pm.IDTipo=tm.IDTipo

go 
-- drop view dbo.vinvExistenciaLote
CREATE VIEW dbo.vinvExistenciaLote
as 
SELECT E.IDBODEGA,B.Descr DescrBodega,E.IDPRODUCTO,P.Descr DescrProductp,E.IDLOTE,L.LoteProveedor,L.LoteInterno,L.FechaVencimiento,L.FechaFabricacion,E.EXISTENCIA
FROM dbo.invEXISTENCIALOTE E
INNER JOIN dbo.invLOTE L ON E.IDLOTE=L.IDLote
INNER JOIN dbo.invBODEGA B ON E.IDBODEGA=B.IDBodega
INNER JOIN dbo.invPRODUCTO P ON E.IDPRODUCTO=P.IDProducto

go 

--ACTUALIZACION DEL INVENTARIO ----------
--DROP PROCEDURE [invUpdateExistenciaBodegaLote]
CREATE PROCEDURE [dbo].[invUpdateExistenciaBodegaLote](@IdBodega INT, @IdProducto INT,@IdLote INT,@Cantidad DECIMAL(28,8))
as 
IF (EXISTS(SELECT * FROM dbo.invEXISTENCIALOTE WHERE IDBODEGA=@IdBodega AND IDPRODUCTO=@IdProducto AND IDLOTE=@IdLote))
	UPDATE dbo.invEXISTENCIALOTE SET EXISTENCIA=EXISTENCIA + @Cantidad 
	WHERE IDBODEGA=@IdBodega and IDPRODUCTO=@IdProducto and IDLOTE=@IdLote
ELSE	
	INSERT INTO dbo.invEXISTENCIALOTE(IDBODEGA,IDPRODUCTO,IDLOTE,EXISTENCIA)
	VALUES(@IdBodega,@IdProducto,@IdLote,@Cantidad)
GO 
--drop procedure [dbo].[invUpdateExistenciaBodega]
CREATE PROCEDURE [dbo].[invUpdateExistenciaBodega](@IdBodega INT, @IdProducto INT,@Cantidad DECIMAL(28,8))
as 
IF (EXISTS(SELECT * FROM dbo.invEXISTENCIABODEGA WHERE IDBODEGA=@IdBodega AND IDPRODUCTO=@IdProducto))
	UPDATE dbo.invEXISTENCIABODEGA SET EXISTENCIA=EXISTENCIA + @Cantidad 
	WHERE IDBODEGA=@IdBodega and IDPRODUCTO=@IdProducto
ELSE	
	INSERT INTO dbo.invEXISTENCIABODEGA(IDBODEGA,IDPRODUCTO,EXISTENCIA)
	VALUES(@IdBodega,@IdProducto,@Cantidad)

GO 



Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEALote foreign key (IDLote)
references dbo.invLOTE (IDLote)

GO 
--drop procedure dbo.invUpdateExistenciaBodegaLinea
CREATE PROCEDURE [dbo].[invUpdateExistenciaBodegaLinea] @IdBodega INT, @IdProducto INT,@IdLote INT = NULL,@Cantidad DECIMAL(28,8), @IdTipoTransaccion INT,@Usuario nvarchar(50)
AS


DECLARE @TRANSACCION as NVARCHAR(20),@FACTOR AS SMALLINT

SELECT @TRANSACCION= Transaccion,@FACTOR=Factor 
	FROM dbo.invTIPOMOVIMIENTO WHERE IDTipo=@IdTipoTransaccion

set @Cantidad = abs(@Cantidad) * @FACTOR

EXEC dbo.invUpdateExistenciaBodegaLote  @IdBodega,@IdProducto,@IdLote,@Cantidad
EXEC dbo.invUpdateExistenciaBodega @IdBodega,@IdProducto,@Cantidad

GO 
--DROP VIEW  DBO.vfafVentaDetalle
CREATE VIEW DBO.vfafVentaDetalle
AS 
SELECT FC.IDBodega,B.Descr DescrBodega, FC.IDFactura, FC.IDCliente,C.Nombre NombreCliente, FC.IDVendedor, FC.Fecha,
       FC.NoPreimpreso, FC.Anulada, FC.BackOrder, FC.IDPedido, FC.EsTeleventa,
       FC.TipoFactura,FL.IDProducto, FL.IDLote, FL.Cantidad, FL.PrecioLocal, FL.PrecioDolar,
       FL.CostoLocal, FL.CostoDolar, FL.TipoCambio, FL.SubTotalLocal,
       FL.SubTotalDolar, FL.SubImpuestoLocal, FL.SubImpuestoDolar, FL.TotalLocal,
       FL.TotalDolar, FL.FactorDevolucion, FL.CantidadDevuelta
  FROM DBO.fafFACTURA FC
INNER JOIN DBO.fafFACTURA_LINEA FL ON FL.IDBodega = FC.IDBodega AND FL.IDFactura = FC.IDFactura AND FL.IDCliente = FC.IDCliente AND FL.IDVendedor = FC.IDVendedor AND FL.Fecha = FC.Fecha
INNER JOIN dbo.ccCLIENTE C ON fc.IDCliente=c.IDCLIENTE 
INNER JOIN dbo.invBODEGA B ON fc.IDBodega=b.IDBodega
WHERE FC.Anulada=0

GO 

--DROP procedure [dbo].[invUpdateMasterExistenciaBodega]
CREATE PROCEDURE [dbo].[invUpdateMasterExistenciaBodega] @Documento AS NVARCHAR(20),@IDPaquete int, @IdTipoTransaccion INT,@Usuario nvarchar(50)
AS 

declare @i int,@iRwCnt int, @IDLote int,  @Cantidad decimal (28,8),@IdProducto AS INT, @IdBodega AS INT	,@CostoLocal AS DECIMAL(28,8),
@CostoDolar AS DECIMAL(28,8)


Create Table #tmpMovimiento( 
	ID int identity(1,1), 
	IDBodega INT, 
	IDProducto INT, 
	IdLote int, 
	Cantidad decimal(28,8) default 0,
	CostoLocal DECIMAL(28,8) DEFAULT 0,
	CostoDolar DECIMAL(28,8) DEFAULT 0,
	IDTipoTransaccion INT)
	create clustered index idx_tmp on #tmpMovimiento(ID) WITH FILLFACTOR = 100

INSERT INTO #tmpMovimiento(IDBodega, IDProducto, IdLote, Cantidad,CostoLocal,CostoDolar,
            IDTipoTransaccion)
SELECT IDBODEGA,IDPRODUCTO,IDLOTE,CANTIDAD,COSTOLOCAL,COSTODOLAR,IDTIPO
FROM dbo.vinvMovimientos  WHERE DOCUMENTO=@Documento AND IDPAQUETE=@IDPaquete AND IDTIPO=@IdTipoTransaccion


set @iRwCnt = @@ROWCOUNT
set @i = 1
set @Cantidad = 0 


while @i <= @iRwCnt 
	begin
		select @IDLote = IdLote, @Cantidad = Cantidad, @IdBodega= IdBodega, @IdProducto= IdProducto ,
				@IdTipoTransaccion=IDTipoTransaccion,@CostoLocal=CostoLocal,@CostoDolar = CostoDolar
		  from #tmpMovimiento where ID = @i
		  
		exec dbo.invUpdateExistenciaBodegaLinea  @IdBodega, @IdProducto,@IDLote, @Cantidad,@IdTipoTransaccion,@Usuario
		
		IF (@IdTipoTransaccion IN (1,3,5,7,10)) --Actualizar el costo del producto si la transaccion es tipo ingreso
		begin
			--Calcular el costo promedio del producto
			UPDATE dbo.invPRODUCTO SET CostoUltLocal = @CostoLocal,CostoUltDolar = @CostoDolar,
					CostoUltPromLocal = CostoUltPromLocal,CostoUltPromDolar = CostoUltPromDolar 
			WHERE IDProducto=@IdProducto
		END
		
		set @i = @i + 1
	end

	
GO  



CREATE PROCEDURE dbo.invGeneraPaqueteFactura(@IDBodega INT,@IDFactura INT,@IDCliente INT,@IDVendedor INT,@Fecha DATETIME,@Usuario AS NVARCHAR(50))
AS 

DECLARE @IDPaquete AS INT,@Documento AS NVARCHAR(20),@Concepto AS NVARCHAR(250)

SET @IDPaquete = 1

--Genera Concepto
SELECT @Concepto= 'Factura: ' + CAST(NoPreimpreso AS NVARCHAR(20)) + ', Bodega=' + DescrBodega + ', Cliente: ' + NombreCliente  
FROM dbo.vfafVentaDetalle 
WHERE IDBodega=@IDBodega AND IDFactura=@IDFactura AND IDCliente=@IDCliente 
AND IDVendedor=@IDVendedor AND Fecha=@Fecha

EXEC dbo.invInsertCabMovimientos @IDPaquete, @Documento OUTPUT,@Fecha,@Concepto,@Usuario,@Usuario

declare @i int,@iRwCnt int, @IDLote int,  @Cantidad decimal (28,8),@IdProducto INT, @IdTipo  INT,@Trasaccion AS INT,
@CostoDolar  DECIMAL(28,8),@CostoLocal DECIMAL(28,8),@PrecioLocal DECIMAL(28,8),@PrecioDolar DECIMAL(28,8)

Create Table #tmpMovimiento( 
	ID int identity(1,1), 
	IDBodega INT, 
	IDProducto INT, 
	IdLote int,
	Documento NVARCHAR(20),
	Fecha DATETIME,
	IdTipo INT,
	Transaccion  NVARCHAR(10),
	Naturaleza NVARCHAR(1), 
	Cantidad decimal(28,8) default 0,
	CostoDolar decimal(28,8) default 0,
	CostoLocal decimal(28,8) default 0,
	PrecioDolar decimal(28,8) default 0,
	PrecioLocal decimal(28,8) default 0)
	create clustered index idx_tmp on #ProductoLote(ID) WITH FILLFACTOR = 100


INSERT INTO  #tmpMovimiento(IDBodega, IDProducto, IdLote, Documento, Fecha,
            IdTipo, Transaccion, Naturaleza, Cantidad, CostoDolar, CostoLocal,
            PrecioDolar, PrecioLocal)
SELECT IDBodega,IDProducto,IDLote,@Documento,Fecha,1 IdTipo,'FAC' Transaccion,'S' Naturaleza,
	Cantidad,CostoDolar,CostoLocal,PrecioDolar,PrecioLocal
  from  dbo.vfafVentaDetalle WHERE IDBodega=@IDBodega AND IDFactura=@IDFactura AND IDCliente=@IDCliente 
AND IDVendedor=@IDVendedor AND Fecha=@Fecha


set @iRwCnt = @@ROWCOUNT
set @i = 1
set @Cantidad = 0 


while @i <= @iRwCnt 
	begin
		select @IDLote = IdLote, @Cantidad = Cantidad, @IdBodega= IdBodega, @IdProducto= IdProducto ,
		@IdTipo=IDTipo
		  from #tmpMovimiento where ID = @i
		  
		  exec dbo.invInsertMovimientos @IDPaquete, @IDBodega, @IDProducto, @IDLote,
			@Documento, @Fecha, @IdTipo, @Trasaccion, 'S', @Cantidad,
			@CostoDolar, @CostoLocal, @PrecioDolar, @PrecioLocal, @Usuario,
			@Usuario
		
		set @i = @i + 1
	end

go
--********************************************
--drop procedure dbo.invGetSugeridoLote
CREATE PROCEDURE dbo.invGetSugeridoLote(@IDBodega INT,@IDProducto INT,@Cantidad DECIMAL(28,8),@CantidadBonificada DECIMAL(28,8))    
AS 
/*SET @Cantidad =25
SET @IDProducto=1
SET @IDBodega=1
*/
set nocount on

declare @iRwCnt INT,@CantidadLote DECIMAL(28,8),@CantidadAsignada DECIMAL(28,8),@Completado BIT
DECLARE @i INT,@IDLote INT

Create Table #Resultado (
IDBodega nvarchar(20), --COLLATE Latin1_General_CI_AS, 
IDProducto nvarchar(20),-- COLLATE Latin1_General_CI_AS, 
IDLote int, Cantidad decimal(28,8) default 0,
IsBonificado BIT DEFAULT 0 
)                            


Create Table #ProductoLote ( ID int identity(1,1), IDBodega INT, 
IDProducto INT , IDLote int, Existencia decimal(28,8) default 0  )

create clustered index idx_tmp on #ProductoLote(ID) WITH FILLFACTOR = 100



/*Existencias*/
insert #ProductoLote (IDBodega, IDProducto, IDLote, Existencia)
SELECT IDBODEGA,IDPRODUCTO,IDLOTE,EXISTENCIA FROM dbo.vinvExistenciaLote
WHERE IDBODEGA=@IDBodega AND IDPRODUCTO=@IDProducto
ORDER BY IDBODEGA,IDPRODUCTO,FechaVencimiento ASC

SET @iRwCnt=@@ROWCOUNT

set @i = 1
set @Completado = 0
set @CantidadLote = 0
set @CantidadAsignada = 0
while @i <= @iRwCnt and @Completado = 0
begin
      select @IDLote = IDLote, @CantidadLote = Existencia from #ProductoLote where ID = @i
      if @Cantidad <= @CantidadLote
      begin
            set @CantidadAsignada = @Cantidad
            insert #Resultado ( IdBodega, IDProducto, IDLote, Cantidad )
            values ( @IDBodega, @IDProducto, @IDLote, @CantidadAsignada )
            set @Completado = 1
      end
      else
      begin
            set @CantidadAsignada = @CantidadLote
            insert #Resultado ( IDBodega, IDProducto, IDLote, Cantidad )
                                         values ( @IDBodega, @IDProducto, @IdLote, @CantidadAsignada )
                                         set @Cantidad = @Cantidad- @CantidadLote
      end
      set @i = @i + 1
END

--Para recorrer los productos bonificados

IF (@CantidadBonificada>0)
BEGIN
	
	
	DBCC CHECKIDENT (#ProductoLote, RESEED, 0)
	DELETE FROM #ProductoLote

	--SELECT * FROM #ProductoLote
    insert #ProductoLote (IDBodega, IDProducto, IDLote, Existencia)
	SELECT A.IDBODEGA,A.IDPRODUCTO,A.IDLOTE,isnull(A.EXISTENCIA,0)- isnull(B.Cantidad,0)
	  FROM dbo.vinvExistenciaLote A
	LEFT JOIN #resultado B ON B.IDBODEGA = A.IDBODEGA AND B.IDPRODUCTO = A.IDPRODUCTO AND B.IDLOTE = A.IDLOTE
	WHERE A.IDBODEGA=@IDBodega AND A.IDPRODUCTO=@IDProducto AND (isnull(A.EXISTENCIA,0)-isnull(B.Cantidad,0))>0
	ORDER BY IDBODEGA,IDPRODUCTO,FechaVencimiento ASC
	
	SET @iRwCnt=@@ROWCOUNT



	set @i = 1
	set @Completado = 0
	set @CantidadLote = 0
	set @CantidadAsignada = 0
	while @i <= @iRwCnt and @Completado = 0
	begin
		  select @IDLote = IDLote, @CantidadLote = Existencia from #ProductoLote where ID = @i
		 
		  if @CantidadBonificada <= @CantidadLote
		  begin
				set @CantidadAsignada = @CantidadBonificada
				insert #Resultado ( IdBodega, IDProducto, IDLote, Cantidad,IsBonificado )
				values ( @IDBodega, @IDProducto, @IDLote, @CantidadAsignada,1 )
				set @Completado = 1
		  end
		  else
		  begin
				set @CantidadAsignada = @CantidadLote
				insert #Resultado ( IDBodega, IDProducto, IDLote, Cantidad,IsBonificado )
											 values ( @IDBodega, @IDProducto, @IdLote, @CantidadAsignada,1)
											 set @CantidadBonificada = @CantidadBonificada- @CantidadLote
		  end
		  set @i = @i + 1
	END
END


SELECT A.IDBodega, A.IDProducto, A.IDLote,L.LoteInterno, L.LoteProveedor,
       L.FechaVencimiento, L.FechaFabricacion, A.Cantidad,A.IsBonificado
  FROM #Resultado A
INNER JOIN dbo.invLOTE L ON L.IDLote = A.IDLote
      
DROP TABLE #Resultado
DROP TABLE #ProductoLote


GO



Create Procedure dbo.invUpdateLote @Operacion nvarchar(1), @IDLote int, @LoteInterno nvarchar(255),@LoteProveedor nvarchar(255), @FechaVencimiento datetime , @FechaProduccion datetime
as
set nocount on
declare @NextCodigo int, @Ok bit
if @Operacion = 'I'
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IDLote) + 1 from dbo.invLOTE ),1)
	while @Ok = 0
	begin
		if not exists (Select IDLote from dbo.invLOTE where IdLote= @NextCodigo)
		begin

			insert dbo.invLOTE (IDLote, LoteInterno, LoteProveedor,
			       FechaVencimiento, FechaFabricacion)
			values (@NextCodigo, @LoteInterno, @LoteProveedor,@FechaVencimiento,@FechaProduccion)
			set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IdLote) + 1 from dbo.invLOTE  ),1)		
		end
	end	
end
if @Operacion = 'U'
begin
 Update dbo.invLOTE SET LoteInterno =  @LoteInterno, LoteProveedor= @LoteProveedor, FechaVencimiento = @FechaVencimiento,FechaFabricacion = @FechaProduccion
 WHERE IDLote = @IDLote
end

if @Operacion = 'D'
begin
	Delete from dbo.invLOTE
	WHERE IDLote = @IDLote
END

GO 

CREATE PROCEDURE dbo.invGetLotes @IDLote INT
AS 
SELECT IDLote, LoteInterno, LoteProveedor, FechaVencimiento, FechaFabricacion
FROM dbo.invLOTE
WHERE (IDLote=@IDLote OR @IDLote=-1)



GO 

--END JULIO

--drop table dbo.GlobalTasaCambioDetalle
Create Table dbo.GlobalTasaCambioDetalle(
	IDTasa nvarchar(20) not null,
	Fecha date not null,
	Monto decimal(28,8) default 0
	)
go	
alter table dbo.GlobalTasaCambioDetalle add constraint pkGlobalTasaCambioDetalle primary key ( IDTasa, Fecha )
go

Alter table  dbo.GlobalTasaCambioDetalle add constraint fkGlobalTasaCambioDetalle foreign key (IDTasa) references [globalCATALOGOS] ( IDCatalogo )
go


Create view dbo.vfafCliente
as 
SELECT    C.IDCLIENTE, C.Nombre, C.RazonSocial, C.Direccion, C.Telefono1, C.Telefono2, 
          C.Telefono3, C.Celular1, C.Celular2, C.email,  C.EsFarmacia, C.NombreFarmacia, 
          C.RUC, C.Propietario, C.IDBodega, B.Descr DescrBodega, 
          C.IDPlazo, P.DESCR DescrPlazo, 
          C.IDMoneda, M.Descr DescrMoneda, 
          C.IDCategoria, R.Descr DescrCatCliente,
          C.IDDepartamento, D.Descr DescrDepto,
          C.IDMunicipio, N.Descr DescrMunicipio,
          C.IDZona, Z.Descr DescrZona,
          C.IDVendedor, V.Nombre NombreVendedor, 
          C.FecUltCredito, C.SALDOLOCAL, C.SALDODOLAR, 
          C.TechoCredito, C.Activo
FROM    dbo.ccCLIENTE C INNER JOIN      dbo.fafVendedor  V 
on C.IDVendedor = V.IDVendedor inner join dbo.invBODEGA B
on C.IDBodega = B.IDBodega inner join dbo.vglobalPlazo P
on C.IDPlazo = P.Codigo inner join dbo.vglobalMoneda M
on C.IDMoneda = M.Codigo INNER JOIN dbo.VGLOBALCATEGORIACLIENTE R
ON C.IDCategoria = R.CODIGO inner join dbo.vglobalDepartamento D
ON C.IDDepartamento = D.CODIGO inner join dbo.vglobalMunicipio N
on C.IDMunicipio = N.codigo inner join dbo.vglobalZona Z
on C.IDZona = Z.codigo
go
	
Create Procedure dbo.fafgetClientes ( @IDCliente int )
as
set nocount on
SELECT [IDCLIENTE] ,[Nombre]  ,[RazonSocial]  ,[Direccion]  ,[Telefono1]  ,[Telefono2]
      ,[Telefono3]  ,[Celular1]  ,[Celular2] ,[email] ,[EsFarmacia] ,[NombreFarmacia]
      ,[RUC],[Propietario] ,[IDBodega] ,[DescrBodega] ,[IDPlazo] ,[DescrPlazo]
      ,[IDMoneda] ,[DescrMoneda] ,[IDCategoria] ,[DescrCatCliente] ,[IDDepartamento] ,[DescrDepto]
      ,[IDMunicipio] ,[DescrMunicipio] ,[IDZona] ,[DescrZona] ,[IDVendedor]
      ,[NombreVendedor] ,[FecUltCredito] ,[SALDOLOCAL] ,[SALDODOLAR] ,[TechoCredito] ,[Activo]
  FROM [dbo].[vfafCliente]
WHERE (IDCLIENTE = @IDCliente  OR @IDCliente = -1)
GO


CREATE PROCEDURE dbo.ccUpdateCliente  @Operacion nvarchar(1),
@IDCLIENTE int  ,
@Nombre nvarchar(255) ,
@RazonSocial nvarchar(255) ,
@Direccion nvarchar(255) ,
@Telefono1 varchar(20) ,
@Telefono2 varchar(20) ,
@Telefono3 varchar(20) ,
@Celular1 varchar(20) ,
@Celular2 varchar(20) ,
@email nvarchar(250) ,
@EsFarmacia bit ,
@NombreFarmacia nvarchar(250) ,
@RUC nvarchar(20) ,
@Propietario nvarchar(255) ,
@IDBodega int  ,
@IDPlazo nvarchar(20)  ,
@IDMoneda nvarchar(20)  ,
@IDCategoria nvarchar(20)  ,
@IDDepartamento nvarchar(20)  ,
@IDMunicipio nvarchar(20)  ,
@IDZona nvarchar(20)  ,
@IDVendedor int  ,
@TechoCredito decimal(28, 8) ,
@Activo bit, @Usuario nvarchar(20)
	
as
declare @NextCodigo int, @Ok bit
set nocount on
if @Operacion = 'I' -- se va a insertar 
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IDCliente) + 1 from dbo.ccCliente  ),1)
	while @Ok = 0
	begin
		if not exists (Select IDCliente from dbo.ccCliente where IDCliente = @NextCodigo)
		begin
			insert [dbo].ccCliente ( [IDCLIENTE]
      ,[Nombre]      ,[RazonSocial] ,[Direccion] ,[Telefono1] ,[Telefono2] ,[Telefono3]  ,[Celular1]
      ,[Celular2] ,[email] ,[EsFarmacia] ,[NombreFarmacia] ,[RUC]  ,[Propietario] ,[IDBodega] ,[IDPlazo]
      ,[IDMoneda] ,[IDCategoria] ,[IDDepartamento] ,[IDMunicipio] ,[IDZona] ,[IDVendedor] ,[TechoCredito],[Activo],
      UserInsert, FechaInsert
			)
			VALUES ( @NextCodigo   ,@Nombre  ,@RazonSocial  ,@Direccion  ,@Telefono1 ,@Telefono2  ,@Telefono3  ,
				@Celular1  ,@Celular2 ,@email  ,@EsFarmacia  ,@NombreFarmacia  ,@RUC  ,
				@Propietario ,@IDBodega   ,@IDPlazo  ,@IDMoneda   ,@IDCategoria   ,@IDDepartamento   ,
				@IDMunicipio   ,@IDZona   ,@IDVendedor, 
				@TechoCredito  ,@Activo, @Usuario, getdate()
				)
		set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IDCliente) + 1 from dbo.ccCliente  ),1)		
		end
	end
end
if @Operacion = 'U' -- se va a Actualizar
begin
	Update [dbo].ccCliente set Nombre = @Nombre  ,RazonSocial = @RazonSocial  ,Direccion = @Direccion  ,
				Telefono1 = @Telefono1 ,Telefono2 = @Telefono2  ,Telefono3 = @Telefono3  ,
				Celular1 = @Celular1  ,Celular2 = @Celular2 ,email = @email  ,EsFarmacia  = @EsFarmacia  ,
				NombreFarmacia = @NombreFarmacia  ,RUC = @RUC  , Propietario = @Propietario ,IDBodega = @IDBodega   ,
				IDPlazo = @IDPlazo  ,IDMoneda = @IDMoneda   ,IDCategoria = @IDCategoria   ,IDDepartamento = @IDDepartamento   ,
				IDMunicipio = @IDMunicipio   ,IDZona = @IDZona   ,IDVendedor = @IDVendedor ,
				TechoCredito = @TechoCredito  ,Activo = @Activo, userUpdate = @Usuario, FechaUpdate = getdate()
	Where IDCliente = @IDCliente
end
if @Operacion = 'D' -- se va a Eliminar
begin
	DELETE FROM [dbo].ccCliente Where [IDCLIENTE] = @IDCliente
end

GO

Create Table dbo.fafEscalaBonificacion ( IDProducto int not null, IDEscala int not null, PorCada decimal(28,4) default 0, 
Bonifica decimal(28,4) default 0
)
go
alter table dbo.fafEscalaBonificacion add constraint pkfafEscalaBonificacion primary key (IDProducto, IDEscala) 
go

alter table dbo.fafEscalaBonificacion add constraint fkfafEscalaBonificacionProducto foreign key  (IDProducto) references dbo.invProducto (IDProducto)
go

-- drop procedure dbo.fafUpdateEscalaBonificacion 
CREATE PROCEDURE dbo.fafUpdateEscalaBonificacion  @Operacion nvarchar(1),
@IDProducto int  ,
@IDEscala int  ,
@PorCada decimal(28,4),
@Bonifica decimal(28,4)
as
declare @NextCodigo int, @Ok bit
set nocount on
if @Operacion = 'I' -- se va a insertar 
begin
	set @Ok = 0
	set @NextCodigo = isnull((Select MAX(IDEscala) + 1 from dbo.fafEscalaBonificacion where IDProducto = @IDProducto ),1)
	while @Ok = 0
	begin
		if not exists (Select IDEscala from dbo.fafEscalaBonificacion where IDProducto = @IDProducto and IDEscala = @NextCodigo)
		begin
			insert dbo.fafEscalaBonificacion ( IDProducto, IDEscala, PorCada, Bonifica )
			VALUES (@IDProducto, @NextCodigo   ,@PorCada  ,@Bonifica  
				)
		set @Ok = 1
		end
		else
		begin
			set @Ok = 0
			set @NextCodigo = isnull((Select MAX(IDEscala) + 1 from dbo.fafEscalaBonificacion where IDProducto = @IDProducto ),1)
		end
	end
end
if @Operacion = 'U' -- se va a Actualizar
begin
	Update dbo.fafEscalaBonificacion set PorCada = @PorCada  ,Bonifica = @Bonifica  
	Where IDPRODUCTO = @IDProducto and IDEscala = @IDEscala 
end
if @Operacion = 'D' -- se va a Eliminar
begin
	DELETE FROM dbo.fafEscalaBonificacion Where IDPRODUCTO = @IDProducto and IDEscala = @IDEscala 
end

GO
-- drop view dbo.vfafEscalaBonificacion
Create View dbo.vfafEscalaBonificacion
as
Select b.IDProducto,  P.Descr, B.IDEscala, B.PorCada, B.Bonifica 
From dbo.fafEscalaBonificacion B inner join dbo.invPRODUCTO P
on B.IDProducto = P.IDProducto 
go

Create Procedure dbo.fafgetEscalaBonificacion @IDProducto int 
as
Set nocount on
Select IDProducto, Descr, IDEscala,  PorCada, Bonifica
From dbo.vfafEscalaBonificacion 
where (IDProducto = @IDProducto or @IDProducto = -1 ) 
order by IDProducto, IDEscala
go

/*
-- PARA INSERTAR VALORES EN LOS CATALOGOS
-- PARA INSERTAR VALOR PROTEGIDO MONEDA

--SELECT * FROM dbo.globalTABLAS 
INSERT [dbo].[globalCATALOGOS] ( [IDCATALOGO] ,[IDTABLA]  ,[CODIGO]  ,[DESCR] ,[ACTIVO], [UsaValor] ,[NombreValor] ,[Valor] ,[Protected] )
VALUES ('0' , 7 ,0  ,'CORDOBA' ,1, 0, 'ND',  0 ,1)
GO
INSERT [dbo].[globalCATALOGOS] ( [IDCATALOGO] ,[IDTABLA]  ,[CODIGO]  ,[DESCR] ,[ACTIVO], [UsaValor] ,[NombreValor] ,[Valor] ,[Protected] )
VALUES ('0' , 7 ,0  ,'DOLAR' ,1, 0, 'ND',  0 ,1)
GO


alter table globalCATALOGOS add CodSistAnterior nvarchar(20)
*/



/*Ultimo Cambio Julio, sin conciliar*/
--DROP TABLE dbo.[invTraslados]
CREATE TABLE [dbo].[invTraslados](
      [IDTraslado] [nvarchar](20) NOT NULL,
      [BodegaOrigen] [int] NOT NULL,
      [BodegaDestino] [int] NOT NULL,
      [IDStatusRecibido] [nvarchar](20) NOT NULL, -- RecibidoParcial, RecibidoTotal, 
      [FechaRemision] [datetime] NULL,
      [FechaEntrada] [datetime] NULL,
      [NumEntrada] [nvarchar](50) NULL,
      [NumSalida] [nvarchar](50) NULL,
      [DocumentoAjuste] [nvarchar](50) NULL,
      Aplicado bit default 0
) ON [PRIMARY]

GO
--DROP TABLE [invDetalleTraslados]
CREATE TABLE [dbo].[invDetalleTraslados](
      [IDTraslado] [nvarchar](20) NOT NULL,
      [BodegaOrigen] [int] NOT NULL,
      [BodegaDestino] [int] NOT NULL,
      [IDProducto] [int] NOT NULL,
      [IDLote] [int] NOT NULL,
      [Cantidad] [decimal](28, 8) default 0, -- esta es la cantidad original del producto en el traslado
      [CantidadRecibida] [decimal](28, 8) default 0, -- esta es la cantidad recibida físicamente
      [Ajuste] [decimal](28, 8) NULL,
      [RecibidoParcial] bit default 0,
      [RecibidoTotal] bit default 0
) ON [PRIMARY]

GO 


--Insercion de Status de Traslados
insert [dbo].[globalTABLAS] (IDTABLA , nombre, Abrev , activo, IDModulo, DescrUsuario )
values (16, 'STATUS_RECIBIDO', 'STATUS_RECIBIDO',1,1000, 'Status Recibido')

GO 

INSERT [dbo].[globalCATALOGOS] ( [IDCATALOGO] ,[IDTABLA]  ,[CODIGO]  ,[DESCR] ,[ACTIVO], [UsaValor] ,[NombreValor] ,[Valor] ,[Protected] )
VALUES (16 , 16 ,1  ,'Recibido Parcial' ,1, 0, 'ND',  0 ,1)

GO 

INSERT [dbo].[globalCATALOGOS] ( [IDCATALOGO] ,[IDTABLA]  ,[CODIGO]  ,[DESCR] ,[ACTIVO], [UsaValor] ,[NombreValor] ,[Valor] ,[Protected] )
VALUES (16 , 16 ,2  ,'Recibido Total' ,1, 0, 'ND',  0 ,1)

GO 

INSERT [dbo].[globalCATALOGOS] ( [IDCATALOGO] ,[IDTABLA]  ,[CODIGO]  ,[DESCR] ,[ACTIVO], [UsaValor] ,[NombreValor] ,[Valor] ,[Protected] )
VALUES (16 , 16 ,2  ,'Pendiente' ,1, 0, 'ND',  0 ,1)

GO 

Create View dbo.vinvStatusRecibido
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'STATUS_RECIBIDO' and activo = 1

GO 

CREATE PROCEDURE dbo.invPreparaDetalleTraslados
AS 
SELECT A.IDProducto,B.Descr DescrProducto,A.IDLote,C.LoteInterno, C.LoteProveedor,
       C.FechaVencimiento, A.Cantidad,a.CantidadRecibida,A.Ajuste
  FROM dbo.invDetalleTraslados a
INNER JOIN dbo.invPRODUCTO B ON B.IDProducto = a.IDProducto
INNER JOIN dbo.invLOTE C ON A.IDLote=C.IDLote
WHERE 1=2

GO 
--DROP procedure invGetCabTraslados
CREATE PROCEDURE dbo.invGetCabTraslados @IdTraslado AS NVARCHAR(20)
AS 
SET NOCOUNT ON

/*SET @IdTraslado=1
SET @NumEntrada ='25260045'*/

SELECT Tra.IDTraslado,  Tra.BodegaOrigen, B1.Descr DescrBodegaOrigen, Tra.BodegaDestino, B2.Descr DescrBodegaDestino, Tra.IDStatusRecibido,E.Descr DescrStatusRecibido,Tra.FechaRemision,
       Tra.FechaEntrada, Tra.NumEntrada, Tra.NumSalida, Tra.DocumentoAjuste, Tra.Aplicado
  FROM dbo.[invTraslados] Tra INNER JOIN dbo.vinvStatusRecibido E ON Tra.IDStatusRecibido=E.Codigo
  INNER JOIN dbo.invBODEGA B1 ON tra.BodegaOrigen=B1.IDBodega
  INNER JOIN dbo.invBODEGA B2 ON tra.BodegaDestino=B2.IDBodega
WHERE (IDTraslado=@IdTraslado OR @IdTraslado='*')


GO
--DROP PROCEDURE invGetCabTrasladosConFiltros
CREATE PROCEDURE dbo.invGetCabTrasladosConFiltros @IdTraslado AS NVARCHAR(20),@Bodega INT,@NumEntrada AS NVARCHAR(50),
										@NumSalida AS NVARCHAR(50),@FechaInicial AS DATETIME,@FechaFinal AS DATETIME,
										@IsEntrada AS INT,@ViewPendienteAplicar AS INT
AS 
SET NOCOUNT ON
/*SET @IdTraslado='*'
SET @NumEntrada ='*'
SET @Bodega=-1
SET @NumSalida='*'
SET @FechaInicial='20140601' 
SET @FechaFinal='20140603'
SET @IsEntrada=1
SET @ViewPendienteAplicar=0*/


SELECT Tra.IDTraslado, Tra.BodegaOrigen, B1.Descr DescrBodegaOrigen, Tra.BodegaDestino, B2.Descr DescrBodegaDestino, Tra.IDStatusRecibido,E.Descr DescrStatusRecibido,Tra.FechaRemision,
       Tra.FechaEntrada, Tra.NumEntrada, Tra.NumSalida, Tra.DocumentoAjuste, Tra.Aplicado
  FROM dbo.[invTraslados] Tra INNER JOIN dbo.vinvStatusRecibido E ON Tra.IDStatusRecibido=E.Codigo
  INNER JOIN dbo.invBODEGA B1 ON tra.BodegaOrigen=B1.IDBodega
  INNER JOIN dbo.invBODEGA B2 ON tra.BodegaDestino=B2.IDBodega
WHERE (IDTraslado=@IdTraslado OR @IdTraslado='*') AND (NumEntrada=@NumEntrada OR @NumEntrada='*')
AND (NumSalida=@NumSalida OR @NumSalida='*') 
AND (@IsEntrada=-1 AND ((FechaRemision BETWEEN @FechaInicial AND @FechaFinal) OR ( FechaEntrada BETWEEN @FechaInicial AND @FechaFinal)) OR (FechaRemision BETWEEN @FechaInicial AND @FechaFinal AND  @IsEntrada=0) OR 
(FechaEntrada BETWEEN @FechaInicial AND @FechaFinal AND @IsEntrada=1) OR (@IsEntrada=1 AND @ViewPendienteAplicar=-1) OR (@IsEntrada=1 AND @ViewPendienteAplicar=1 AND FechaEntrada='1980-01-01 00:00:00.000')) 
AND (@ViewPendienteAplicar=-1 OR (@ViewPendienteAplicar=1 AND FechaEntrada='1980-01-01 00:00:00.000') OR @ViewPendienteAplicar=0)
AND ( @Bodega=-1 OR (@IsEntrada=1 AND Tra.BodegaDestino = @Bodega) OR (@IsEntrada=0 AND tra.BodegaOrigen=@Bodega) OR (@IsEntrada=-1 AND( tra.BodegaOrigen=@Bodega OR tra.BodegaDestino=@Bodega)))

GO 

Create Procedure dbo.invUpdateCabTraslados(@Operacion nvarchar(1), @IDTraslado NVARCHAR(20), @BodegaOrigen AS INT, @BodegaDestino AS INT, @IDStatusRecibido AS NVARCHAR(20), 
											@FechaRemision AS DATETIME,@FechaEntrada AS DATETIME,@NumEntrada AS NVARCHAR(50),@NumSalida AS NVARCHAR(50),@DocumentoAjuste AS NVARCHAR(20),
											@Aplicado AS BIT)
as
set nocount on
DECLARE  @Ok bit
if @Operacion = 'I'
begin
	set @Ok = 0
	
	exec dbo.invGetNextConsecutivoPaquete 7,@IDTraslado OUTPUT
	while @Ok = 0
	begin
		if not exists (SELECT IDTraslado from [dbo].[invTraslados] WHERE IDTraslado= @IDTraslado)
		begin
			SET @IDStatusRecibido='16-3'
			insert dbo.[invTraslados](IDTraslado, BodegaOrigen, BodegaDestino,
			       IDStatusRecibido, FechaRemision, FechaEntrada, NumEntrada,
			       NumSalida, DocumentoAjuste, Aplicado)
			values (@IDTraslado,@BodegaOrigen,@BodegaDestino,@IDStatusRecibido,@FechaRemision,
					@FechaEntrada,@NumEntrada,@NumSalida,@DocumentoAjuste,@Aplicado)
			set @Ok = 1
			SELECT @IDTraslado IDTraslado
		end
		else
		begin
			set @Ok = 0
			exec dbo.invGetNextConsecutivoPaquete 7,@IDTraslado OUTPUT		
		end
	end	
end
if @Operacion = 'U'
begin
 Update dbo.[invTraslados] SET IDStatusRecibido = @IDStatusRecibido,NumEntrada = @NumEntrada,NumSalida = @NumSalida,DocumentoAjuste= @DocumentoAjuste,
								Aplicado = @Aplicado,FechaRemision = @FechaRemision, FechaEntrada = @FechaEntrada
 WHERE IDTraslado = @IDTraslado
end

if @Operacion = 'D'
begin
	Delete from dbo.invTraslados 
	WHERE IDTraslado=@IDTraslado
end

GO


CREATE PROCEDURE dbo.invGetDetalleTraslados @IdTraslado AS NVARCHAR(20)
AS 
SET NOCOUNT ON

SELECT D.IDTraslado, D.BodegaOrigen, D.BodegaDestino, D.IDProducto,p.Descr DescrProducto, D.IDLote,L.LoteInterno,
       L.LoteProveedor, L.FechaVencimiento,D.Cantidad, D.CantidadRecibida, D.Ajuste, D.RecibidoParcial,
       D.RecibidoTotal
  FROM dbo.invDetalleTraslados D INNER JOIN dbo.invPRODUCTO P ON D.IdProducto = P.IDProducto
  INNER JOIN dbo.invLOTE L ON D.IdLote =L.IDLote 
WHERE (IDTraslado=@IdTraslado OR @IdTraslado='*') 

GO 


Create Procedure dbo.invUpdateDetalleTraslados @Operacion nvarchar(1), @IDTraslado NVARCHAR(20), @BodegaOrigen INT, @BodegaDestino INT, @IdProducto INT, 
											@IdLote int,@Cantidad DECIMAL(28,8),@CantidadRecibida  DECIMAL(28,8),@Ajuste DECIMAL(28,8),@RecibidoParcial BIT,
											@RecibidoTotal BIT
as
set nocount on

if @Operacion = 'I'
BEGIN
	insert dbo.invDetalleTraslados(IDTraslado, BodegaOrigen, BodegaDestino,
	       IDProducto, IDLote, Cantidad, CantidadRecibida, Ajuste, RecibidoParcial,
	       RecibidoTotal)
	values ( @IDTraslado,@BodegaOrigen,@BodegaDestino,@IdProducto,@IdLote,@Cantidad,@CantidadRecibida,@Ajuste,@RecibidoParcial,@RecibidoTotal)
end
if @Operacion = 'U'
begin
 Update dbo.invDetalleTraslados SET Cantidad = @Cantidad,
     CantidadRecibida = @CantidadRecibida,
     Ajuste = @Ajuste,
     RecibidoParcial = @RecibidoParcial,
     RecibidoTotal = @RecibidoTotal
 WHERE IDTraslado = @IDTraslado AND IDProducto=@IdProducto AND IDLote=@IdLote
end

if @Operacion = 'D'
begin
	Delete from dbo.invDetalleTraslados 
	WHERE IDTraslado=@IDTraslado AND (IDProducto=@IdProducto OR @IdProducto=-1) AND (IDLote=@IdLote OR @IdLote=-1)
END

IF EXISTS (SELECT IDProducto
             FROM dbo.invDetalleTraslados WHERE IDTraslado=@IDTraslado AND RecibidoParcial=1 AND @Operacion='U') 
	UPDATE dbo.invTraslados SET IDStatusRecibido = '16-1' WHERE IDTraslado=@IDTraslado AND Aplicado = 1
ELSE 
	UPDATE dbo.invTraslados SET IDStatusRecibido = '16-2' WHERE IDTraslado=@IDTraslado AND Aplicado = 1
GO

--drop procedure invGeneraDetalleMovimientoTraslado
CREATE PROCEDURE [dbo].[invGeneraDetalleMovimientoTraslado] @IDTraslado NVARCHAR(20),@IsSalida BIT,@UserInsert NVARCHAR(50),@UserUpdate NVARCHAR(50)
AS 
SET NOCOUNT ON

DECLARE @IDTransac AS INT
/*SET @IDTraslado='TRS000000000002'
SET @IsSalida=0
SET @UserInsert='jespinoza'
SET @UserUpdate='jespinoza'*/

     
	--9	 TRS	Traslado Salida	S	-1	1
	--10 TRE	Traslado Entrada	E	1	1
	
SELECT 7 IDPaquete, case WHEN (@IsSalida=1) then A.BodegaOrigen ELSE a.BodegaDestino END  Bodega,A.IDProducto,A.IDLote,A.IDTraslado,c.FECHA Fecha,
		case when (@IsSalida =1) then 9 ELSE 10 END IDTipo, CASE WHEN (@IsSalida=1) THEN 'TRS' ELSE 'TRE' END Transaccion, CASE WHEN (@IsSalida=1) THEN 'S' ELSE 'E' END  Naturaleza,
		case when (@IsSalida =1) then A.Cantidad ELSE a.CantidadRecibida END Cantidad,B.CostoUltPromLocal,B.CostoUltPromDolar,0 PrecioDolar,
		GETDATE() FechaInsert,GETDATE() FechaUpdate INTO #tmpDetalle
FROM dbo.invDetalleTraslados A
INNER JOIN dbo.invPRODUCTO B ON B.IDProducto = A.IDProducto
INNER JOIN dbo.invCABMOVIMIENTOS C ON A.IDTraslado=C.DOCUMENTO
WHERE IDTraslado=@IDTraslado

select @IDTransac= IDTipo FROM #tmpDetalle

INSERT INTO dbo.invMOVIMIENTOS(IDPAQUETE, IDBODEGA, IDPRODUCTO, IDLOTE, DOCUMENTO,
            FECHA, IDTIPO, TRANSACCION, NATURALEZA, CANTIDAD, COSTOLOCAL,
            COSTODOLAR, PRECIOLOCAL, PRECIODOLAR, UserInsert, UserUpdate,
            FechaInsert, FechaUpdate)  
              
SELECT IdPaquete,Bodega,IDProducto,IDLote,IDTraslado,Fecha,IDTipo,Transaccion,Naturaleza,Cantidad,CostoUltPromLocal,CostoUltPromDolar, 
		0,PrecioDolar,@UserInsert,@UserUpdate,FechaInsert,FechaUpdate  FROM #tmpDetalle


DROP TABLE #tmpDetalle

EXEC DBO.[invUpdateMasterExistenciaBodega] @IDTraslado,7 , @IDTransac,@UserInsert

GO 


--drop procedure invGeneraAjusteByTraslado
create procedure dbo.invGeneraAjusteByTraslado @Documento NVARCHAR(20), @Usuario AS NVARCHAR(20)
AS 
SET NOCOUNT ON

 

DECLARE @IDPaquete AS INT
DECLARE @IDTipoMov AS INT
DECLARE @Transaccion AS NVARCHAR(20)
DECLARE @Naturaleza AS NVARCHAR(1)
DECLARE @DocumentoAjuste AS NVARCHAR(20)
DECLARE @Fecha AS DATETIME, @Concepto  AS NVARCHAR(255), @Referencia AS NVARCHAR(20),@ActualizaConsecutivo AS BIT

--SET @Documento='TRS000000000006'
SET @IDPaquete=3 --Paquete de Ajuste
SET @IDTipoMov=4 --Ajuste de Salida
SET @Transaccion='AJS'
SET @Naturaleza='S'
SET @ActualizaConsecutivo=1
SET @Referencia=@Documento


SELECT @Fecha=FechaEntrada,@Concepto='Ajuste por faltantes en Traslado ' + @Documento + ', de bodega origen: ' + B.Descr + ' y bodega Destino: ' + C.Descr + ' NumEntrada: ' + a.NumEntrada + ' NumSalida: ' + A.numsalida
FROM dbo.invTraslados A
INNER JOIN dbo.invBODEGA B on A.BodegaOrigen=B.IDBodega
INNER JOIN dbo.invBODEGA C ON A.BodegaDestino=C.IDBodega
WHERE IDTraslado=@Documento

--TABLA TEMPORAL CON UNA COLUMNA
DECLARE @t TABLE ( resultado VARCHAR(20) )

INSERT INTO @t 
EXEC dbo.invInsertCabMovimientos @IDPaquete,@DocumentoAjuste,
                                                @Fecha, @Concepto, @Referencia,
                                                @Usuario, @Usuario,
                                                @ActualizaConsecutivo  

SELECT @DocumentoAjuste = resultado FROM @t

Declare @iRowCount int, @iCounter int, @Articulo nvarchar(20)

SELECT BodegaDestino, IDProducto, IDLote,Ajuste INTO #tmpAjustes FROM dbo.invDetalleTraslados
WHERE IDTraslado=@Documento AND Ajuste>0

set @iRowCount  = @@RowCount
Alter table #tmpAjustes add ID int identity(1,1)
--select * from #fmlDetalleFormula where IdFormula = @IDFormula
Create clustered index _tmpAjustes on #tmpAjustes (ID) with fillfactor = 100

set @iCounter = 1

DECLARE @Bodega AS INT
DECLARE @IDProducto AS INT
DECLARE @IDLote AS INT
DECLARE @Cantidad AS DECIMAL(28,8)
DECLARE @CostoDolar AS DECIMAL(28,8)
DECLARE @CostoLocal AS DECIMAL(28,8)

WHILE (@iCounter <= @iRowCount )
BEGIN -- 
	SELECT @Bodega=BodegaDestino, @IDProducto= IDProducto,@IDLote=IDLote, @Cantidad = Ajuste from #tmpAjustes where ID = @iCounter 
	SELECT @CostoLocal =CostoUltPromLocal,@CostoDolar =CostoUltPromDolar
	FROM dbo.invPRODUCTO WHERE IDProducto=@IDProducto
	--Insertar el detalle del ajuste
	EXEC dbo.invInsertMovimientos @IDPaquete,@Bodega,@IDProducto,@IDLote,@DocumentoAjuste,@Fecha,@IDTipoMov,@Transaccion,@Naturaleza,@Cantidad,@CostoDolar,@CostoLocal,0,0,@Usuario,@Usuario
	SET @iCounter = @iCounter + 1
END  


--Actualizar el trasaldo con el documento de ajuste
UPDATE dbo.invTraslados SET DocumentoAjuste = @DocumentoAjuste WHERE IDTraslado=@Documento

GO 


CREATE PROCEDURE dbo.invGetExistenciaBodegaLote(@IdBodega AS INT,@IdProducto AS INT)
AS 
SELECT E.IDBODEGA, B.Descr DescrBodega, E.IDPRODUCTO, E.IDLOTE, E.EXISTENCIA
  FROM dbo.invEXISTENCIALOTE E
INNER JOIN dbo.invLOTE L ON L.IDLote = E.IDLOTE
INNER JOIN dbo.invBODEGA B ON E.IDBODEGA=B.IDBodega
INNER JOIN dbo.invPRODUCTO P ON E.IDPRODUCTO=P.IDProducto
WHERE (E.IDBODEGA=@IdBodega OR @IDBodega=-1) 
AND (E.IDPRODUCTO=@IdProducto OR @IdProducto=-1) 


GO 
--drop procedure dbo.invGetBodegaByUsuario
CREATE PROCEDURE dbo.invGetBodegaByUsuario(@Usuario AS NVARCHAR(50))
AS 
SELECT a.IDBodega, a.Descr DescrBodega, a.Activo, a.PreFactura, a.ConsecFactura,
       a.ConsecPedido, B.Usuario, B.Factura, B.ConsultaInv
  FROM dbo.invBODEGA a
INNER JOIN dbo.invbodegaUsuario B ON B.IDBodega = a.IDBodega
WHERE B.Usuario=@Usuario AND a.Activo=1

GO 


CREATE PROCEDURE dbo.invGetLineas(@IDLinea AS NVARCHAR(20))
AS 
SELECT Codigo, descr, activo, IDModulo FROM dbo.vinvClasificacion1
WHERE Codigo=@IDLinea OR @IDLinea='*'

GO 

CREATE PROCEDURE dbo.invGetFamilias(@IDFamilia AS NVARCHAR(20))
AS 
SELECT Codigo, descr, activo, IDModulo FROM dbo.vinvClasificacion2
WHERE Codigo=@IDFamilia  OR @IDFamilia='*'

GO 

CREATE PROCEDURE dbo.invGetSubFamilias(@IDSubFamilia AS NVARCHAR(20))
AS 
SELECT Codigo, descr, activo, IDModulo FROM dbo.vinvClasificacion3
WHERE Codigo=@IDSubFamilia  OR @IDSubFamilia ='*'

GO 




