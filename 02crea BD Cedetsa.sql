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
	[Factura] [bit] default 0
 CONSTRAINT [PK_invBodega] PRIMARY KEY CLUSTERED 
(
	[IDBodega] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

Create Procedure dbo.invGetBodegas @IDBodega int
as
set nocount on
Select IDBodega, Descr DescrBodega, Activo, Factura 
From dbo.invBODEGA 
where (IDBodega = @IDBodega or @IDBodega = -1)
go
--use Ced2 drop procedure dbo.invUpdateBodega 
Create Procedure dbo.invUpdateBodega @Operacion nvarchar(1), @IDBodega int, @Descr nvarchar(255), @Activo bit, @Factura bit
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

			insert dbo.invBODEGA  (IDBodega, Descr, Activo, Factura)
			values (@NextCodigo, @Descr, @Activo, @Factura)
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

CREATE TABLE [dbo].[globalCATALOGOS](
	[IDCATALOGO] [nvarchar](20) NOT NULL,
	[IDTABLA] [int] NOT NULL,
	[CODIGO] [int] NOT NULL,
	[DESCR] [nvarchar](250) NULL,
	[ACTIVO] [bit] default 1,
	UsaValor bit default 0,
	NombreValor nvarchar(100) default 'ND',
	Valor decimal(8,2) default 0,
	
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
	
	INSERT DBO.globalCATALOGOS (IDCATALOGO, IDTABLA, CODIGO, DESCR, ACTIVO, UsaValor, NombreValor, Valor)
	SELECT CAST(I.IDTABLA AS nvarchar(20)) + CAST (@NEXT AS NVARCHAR(20)),  I.IDTABLA, @NEXT, I.DESCR, I.ACTIVO, I.UsaValor, I.NombreValor, I.Valor
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



Create Procedure globalGetCatalogos @IDTabla int
as
set nocount on
Select C.IDCATALOGO, C.IDTABLA, T.nombre,  C.CODIGO, C.DESCR, C.Activo, C.UsaValor, C.NombreValor, C.Valor 
from DBO.globalCATALOGOS C inner join dbo.globalTABLAS T
on C.IDTABLA = T.IDTABLA 
Where (C.idtabla = @IDTabla or @IDTabla = -1 )
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
CodigoBarra nvarchar(128)
 
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
--DROP VIEW dbo.vinvProducto
Create View dbo.vinvProducto
as
Select IDProducto, Descr, Impuesto, case when EsMuestra = 1 then 'SI' ELSE 'NO' end EsMuestra,
case when EsControlado = 1 then 'SI' ELSE 'NO' end EsControlado,
case when EsEtico = 1 then 'SI' ELSE 'NO' end EsEtico,
Clasificacion1 Linea, Clasificacion2 Familia, Clasificacion3 SubFamilia,
case when BajaPrecioDistribuidor = 1 then 'SI' ELSE 'NO' end BajaPrecioDistribuidor,
case when BajaPrecioProveedor  = 1 then 'SI' ELSE 'NO' end BajaPrecioProveedor,
IDProveedor, CostoUltLocal, CostoUltDolar, CostoUltPromLocal, CostoUltPromDolar, 
PrecioPublicoLocal, PrecioFarmaciaLocal , PrecioCIFLocal, PrecioFOBLocal, 
IDPresentacion, PorcDescAlzaProveedor, UserInsert, FechaInsert,  UserUpdate, FechaUpdate,
CodigoBarra, case when Activo = 1 then 'SI' ELSE 'NO' end Activo
From invProducto
go
-- Drop procedure dbo.invUpdateProducto
CREATE PROCEDURE dbo.invUpdateProducto  @Operacion nvarchar(1),
	@IDProducto int  ,@Descr nvarchar(250) ,@Impuesto nvarchar(20) ,@EsMuestra bit ,@EsControlado bit ,@Clasificacion1 nvarchar(20) ,
	@Clasificacion2 nvarchar(20) ,@Clasificacion3 nvarchar(20) ,@EsEtico bit ,@BajaPrecioDistribuidor bit ,	@IDProveedor int ,
	@CostoUltLocal decimal(28, 8) ,	@CostoUltDolar decimal(28, 8) ,	@CostoUltPromLocal decimal(28, 8) ,	@CostoUltPromDolar decimal(28, 8) ,
	@PrecioPublicoLocal decimal(28, 8) ,	@PrecioFarmaciaLocal decimal(28, 8) ,	@PrecioCIFLocal decimal(28, 8) ,
	@PrecioFOBLocal decimal(28, 8) ,	@IDPresentacion nvarchar(20) ,	@BajaPrecioProveedor bit ,	@PorcDescAlzaProveedor decimal(8, 2) ,
	@UserInsert nvarchar(20) ,	@UserUpdate nvarchar(20) ,	@Activo bit ,	@CodigoBarra nvarchar(128)
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
				  ,[UserUpdate] ,[Activo] ,[CodigoBarra]
			)
			VALUES ( @NextCodigo   ,@Descr ,@Impuesto  ,@EsMuestra  ,@EsControlado  ,@Clasificacion1  ,	@Clasificacion2 ,
				@Clasificacion3  ,	@EsEtico  ,	@BajaPrecioDistribuidor  ,	@IDProveedor  ,	@CostoUltLocal  ,	@CostoUltDolar  ,
				@CostoUltPromLocal  ,	@CostoUltPromDolar  ,	@PrecioPublicoLocal  ,	@PrecioFarmaciaLocal  ,	@PrecioCIFLocal  ,
				@PrecioFOBLocal  ,	@IDPresentacion  ,	@BajaPrecioProveedor  ,	@PorcDescAlzaProveedor  ,	@UserInsert  ,
				@UserUpdate  ,	@Activo  ,	@CodigoBarra 
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
		Activo = @Activo  ,	CodigoBarra = @CodigoBarra 
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

--DROP VIEW dbo.vinvClasificacion3
Create View dbo.vinvTipoFactura
as
SELECT IDCatalogo Codigo, descr,  activo, IDModulo
FROM dbo.vglobalCatalogos 
WHERE Tabla = 'TIPO_FACTURA' and activo = 1
go

--select * from dbo.vinvClasificacion1 exec dbo.invGetProductos -1
-- DROP PROCEDURE dbo.invGetProductos 
Create Procedure dbo.invGetProductos @IDProducto int
as 
set nocount on 

Select P.IDProducto ,P.Descr  ,P.Impuesto  , I.Descr DescrImpuesto, P.EsMuestra ,P.EsControlado ,
		P.Linea  , C1.Descr DescrClasif1, P.Familia, C2.Descr DescrClasif2
		,P.SubFamilia , C3.Descr DescrClasif3, P.EsEtico  ,P.BajaPrecioDistribuidor  ,P.IDProveedor , R.Nombre ,P.CostoUltLocal  ,P.CostoUltDolar
		  , P.CostoUltPromLocal ,P.CostoUltPromDolar ,P.PrecioPublicoLocal ,P.PrecioFarmaciaLocal ,P.PrecioCIFLocal
		  ,P.PrecioFOBLocal ,P.IDPresentacion , S.Descr DescrPresentacion, P.BajaPrecioProveedor ,P.PorcDescAlzaProveedor ,P.UserInsert
		  ,P.UserUpdate ,P.Activo ,P.CodigoBarra
From dbo.vinvProducto P left join dbo.vinvImpuesto I 
on p.Impuesto = I.Codigo left join dbo.vinvClasificacion1 C1
on p.Linea=C1.Codigo left join dbo.vinvClasificacion2 C2
on p.Familia=C2.Codigo left join dbo.vinvClasificacion3 C3
on p.SubFamilia=C3.Codigo left join dbo.cpProveedor R
on p.IDProveedor = R.IDProveedor left join dbo.vinvPresentacion S
on P.IDPresentacion = S.codigo
where (@IDProducto = -1 or IDProducto = @IDProducto)
go

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

-- para modulo Cuentas por Cobrar
-- drop table ccCLIENTE 
Create Table ccCLIENTE ( 
	CodCliente int not null, Nombre nvarchar(255), RazonSocial nvarchar(255), Direccion nvarchar(255),
	TechoCredito decimal(28,8) default 0, PlazoCredito int default 0, 
	Activo bit default 1, Moneda nvarchar(3), FecUltCredito date,
	SALDOLOCAL decimal(28,8) default 0,
	SALDODOLAR decimal(28,8) default 0)
go
alter table ccCLIENTE ADD CONSTRAINT pkCliente primary key (CodCliente )
go
--drop table ccDocumentos
Create Table ccDocumentos ( CodCliente int not null,  Documento nvarchar(20) not null, Tipo nvarchar(3) not null, Fecha date not null, 
MontoLocal decimal(28,8) default 0, MontoDolar decimal(28,8) default 0,
Moneda nvarchar(10), FechaVencimiento date, TipoCambio decimal(28,8) default 0,
EsDebito bit default 0, Aprobado bit default 0
)
go
alter table ccDocumentos add constraint pkccDocumentos primary key clustered (CodCliente, Documento, Tipo, Fecha)
go

--drop table ccDocumentosAplicacion
Create Table ccDocumentosAplicacion ( 
CodCliente int not null,
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
(CodCliente, DocDebito, TipoDebito, DocCredito, TipoCredito)
go



-- Para gragar una Aplicación

--drop procedure dbo.ccUpdateAplicaciones
Create Procedure dbo.ccUpdateAplicaciones
@CodCliente int, @DocDebito nvarchar(20), @TipoDebito nvarchar(3),
@DocCredito nvarchar(20), @TipoCredito nvarchar(3), 
@ValorCreditoDol decimal(28,8), @ValorCreditoLoc decimal(28,8), @Modalidad nvarchar(1),
@FechaCredito date, @TipoCambio decimal(28,8), @MonedaDebito nvarchar(3), @MonedaCredito nvarchar(3)
-- @Modalidad = 'I' Insertar , 'D' Delete 'R' Revertir Aplicacion
as
set nocount on
/*
set @CodCliente =1
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
			where CodCliente = @CodCliente and DocDebito = @DocDebito  and TipoDebito = @TipoDebito  ) 
BEGIN
	Select  top 1 @SaldoInicDebitoDol = MontoDolar, @SaldoFinalDebitoDol = MontoDolar - @ValorCreditoDol ,
	@SaldoInicDebitoLoc = MontoLocal, @SaldoFinalDebitoLoc = MontoLocal - @ValorCreditoLoc 
	From ccDocumentos
	Where CodCliente = @CodCliente and Documento = @DocDebito  and Tipo = @TipoDebito  
END
ELSE -- HAY AL MENOS UNA APLICACION A ESE DEBITO
BEGIN

	Select  top 1 @SaldoInicDebitoDol = SaldoFinalDebitoDol, @SaldoFinalDebitoDol = @SaldoInicDebitoDol - @ValorCreditoDol ,
	@SaldoInicDebitoLoc = SaldoFinalDebitoLoc, @SaldoFinalDebitoLoc = @SaldoInicDebitoLoc - @ValorCreditoLoc 
	From ccDocumentosAplicacion
	Where CodCliente = @CodCliente and DocDebito = @DocDebito  and TipoDebito = @TipoDebito  
	and FechaAplicacion <= @FechaCredito 
	order by CodCliente, DocDebito, FechaAplicacion desc
	
END			

-- Actualizar el saldo del Doc tipo CREDITOS

if not Exists (Select DocCredito -- SI NO EXISTE NINGUNA APLICACION con ESE CREDITO
			From ccDocumentosAplicacion 
			where CodCliente = @CodCliente and DocCredito = @DocCredito  and TipoCredito = @TipoCredito  ) 
BEGIN
	Select Top 1 @SaldoInicCreditoDol = MontoDolar, @SaldoFinalCreditoDol = MontoDolar - @ValorCreditoDol ,
	@SaldoInicCreditoLoc = MontoLocal, @SaldoFinalCreditoLoc = MontoLocal - @ValorCreditoLoc 
	From ccDocumentos
	Where CodCliente = @CodCliente and Documento = @DocCredito  and Tipo = @TipoCredito  
END
ELSE -- HAY AL MENOS UNA APLICACION CON ESE CREDITO
BEGIN
	Select  top 1 @SaldoInicCreditoDol = SaldoFinalCreditoDol, @SaldoFinalCreditoDol = SaldoFinalCreditoDol - @ValorCreditoDol ,
	@SaldoInicCreditoLoc = SaldoFinalCreditoLoc, @SaldoFinalCreditoLoc = SaldoFinalCreditoLoc - @ValorCreditoLoc 
	From ccDocumentosAplicacion
	Where CodCliente = @CodCliente and DocCredito = @DocCredito  and TipoDebito = @TipoCredito
	and FechaAplicacion <= @FechaCredito 
	order by CodCliente, DocCredito, FechaAplicacion desc
	
END			

if @Modalidad = 'I'
begin
	insert ccDocumentosAplicacion(CodCliente, TipoDebito, DocDebito, TipoCredito, DocCredito,FechaAplicacion, 
	MontoDolar, MontoLocal,	
	SaldoInicDebitoDol, SaldoFinalDebitoDol, SaldoInicDebitoLoc, SaldoFinalDebitoLoc,
	SaldoInicCreditoDol, SaldoFinalCreditoDol, SaldoInicCreditoLoc,	SaldoFinalCreditoLoc,	
	TipoCambio, MonedaDebito, MonedaCredito )
	values (
		@CodCliente, @TipoDebito, @DocDebito , @TipoCredito, @DocCredito,  @FechaCredito,
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

--drop table dbo.fafPEDIDO
CREATE TABLE dbo.fafPEDIDO
(
	IDPedido int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	Aprobado bit default 0,
	BackOrder bit default 0
)
go
Alter Table dbo.fafPEDIDO add constraint pkfafPEDIDO primary key clustered ( IDPedido, IDCliente, 
		IDVendedor, Fecha )
go		
Alter Table dbo.fafPedido add constraint fkfafPedidoCliente foreign key (IDCliente) references dbo.ccCliente (CodCliente)
go

Alter Table dbo.fafPedido add constraint fkfafPedidoVendedor foreign key (IDVendedor) references dbo.fafVendedor (IDVendedor)
go


Create Table dbo.fafPEDIDO_LINEA		
(
	IDPedido int not null,
	IDCliente int not null,
	IDVendedor int not null,	
	Fecha datetime not null,
	IDProducto int not null,
	CantidadPedida decimal(28,8) default 0,
	CantidadFacturada decimal(28,8) default 0,
	CantidadNoFacturada decimal(28,8) default 0
)
go
Alter Table dbo.fafPEDIDO_LINEA add constraint pkfafPEDIDO_LINEA primary key clustered ( IDPedido, IDCliente, 
		IDVendedor, Fecha, IDProducto )
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEA foreign key ( IDPedido, IDCliente, 
		IDVendedor, Fecha ) references dbo.fafPedido ( IDPedido, IDCliente, 
		IDVendedor, Fecha )
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEAVendedor foreign key (IDVendedor) references dbo.fafVendedor (IDVendedor)
go


Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEACliente foreign key (IDCliente) references dbo.ccCliente (CodCliente)
go

Alter Table dbo.fafPEDIDO_LINEA add constraint fkfafPEDIDO_LINEAProducto foreign key (IDProducto) references dbo.invProducto (IDProducto)
go

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

Alter Table dbo.fafFactura add Constraint fkfafFacturaCliente foreign key (IDCliente) references dbo.ccCliente( CodCliente )
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
		IDVendedor, Fecha, IDProducto )
go	

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEA foreign key ( IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha ) references dbo.fafFACTURA (IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha)
go		
Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEABodega foreign key (IDBodega)
references dbo.invBodega (IDBodega)
go

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEACliente foreign key (IDCliente)
references dbo.ccCLIENTE (CodCliente)
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

GO


/*
--*********
select *
from ccDocumentosAplicacion 

select *
Delete from ccDocumentosAplicacion 
where TipoDebito = 'FAC' and DocCredito = 'REC1214'


SElect @SaldoInicDebitoDol SaldoInicDebitoDol, @SaldoFinalDebitoDol SaldoFinalDebitoDol,
@SaldoInicDebitoLoc SaldoInicDebitoLoc, @SaldoFinalDebitoLoc SaldoFinalDebitoLoc


exec dbo.fafUpdateVendedor 'I',2,'g','101',1
EXEC dbo.fafUpdateVendedor 'I',4,'DFGASFDAS','101',1
SELECT * FROM DBO.FAFVENDEDOR
*/

