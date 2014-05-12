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

SELECT * FROM dbo.invTIPOMOVIMIENTO
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (8,11,'FIS')
GO 
INSERT [dbo].[invPAQUETETIPOMOV] (IDPaquete, IDTipo, Transaccion )
VALUES (8,12,'FIS')
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
CREATE PROCEDURE [dbo].[invInsertCabMovimientos]  @IDPaquete int,
	@Documento nvarchar(20) OUTPUT,@Fecha DATETIME,@Concepto NVARCHAR(255),@Referencia  nvarchar(20),@UserInsert AS NVARCHAR(20),
	@UserUpdate  NVARCHAR(20)
	
as

set nocount ON

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


CREATE PROCEDURE dbo.invGetDetalleMovimiento(@IDPaquete AS INT,@Documento AS NVARCHAR(20))
AS 


SELECT IDBODEGA,DescrBodega, IDPRODUCTO, DescrProducto, IDLOTE, LoteInterno,
       FechaVencimiento, FechaFabricacion, DOCUMENTO, FECHA, IDTIPO, DescrTipo,
       TRANSACCION, NATURALEZA, CANTIDAD, COSTOLOCAL, COSTODOLAR, PRECIOLOCAL,
       PRECIODOLAR, UserInsert
  FROM dbo.vinvMovimientos
WHERE IDPAQUETE=@IDPaquete AND DOCUMENTO=@Documento 

GO 


CREATE VIEW dbo.vinvPaqueteTipoMovimiento
AS 
SELECT p.IDPaquete, p.PAQUETE, p.Descr DescrPaquete,tm.IDTipo, tm.Transaccion, tm.Descr DescrTipo,
       tm.Naturaleza, tm.Factor, tm.[ReadOnly] 
FROM dbo.invPAQUETE p
INNER JOIN dbo.invPAQUETETIPOMOV pm ON pm.IDPaquete = p.IDPaquete
INNER JOIN dbo.invTIPOMOVIMIENTO tm ON pm.IDTipo=tm.IDTipo

go 

CREATE VIEW dbo.vinvExistenciaLote
as 
SELECT E.IDBODEGA,B.Descr DescrBodega,E.IDPRODUCTO,P.Descr DescrProductp,E.IDLOTE,L.LoteInterno,L.FechaVencimiento,L.FechaFabricacion,E.EXISTENCIA
FROM dbo.invEXISTENCIALOTE E
INNER JOIN dbo.invLOTE L ON E.IDLOTE=L.IDLote
INNER JOIN dbo.invBODEGA B ON E.IDBODEGA=B.IDBodega
INNER JOIN dbo.invPRODUCTO P ON E.IDPRODUCTO=P.IDProducto

go 

--ACTUALIZACION DEL INVENTARIO ----------

CREATE PROCEDURE dbo.invUpdateExistenciaBodegaLote(@IdBodega INT, @IdProducto INT,@IdLote INT,@Cantidad DECIMAL(28,8))
as 
IF (EXISTS(SELECT * FROM dbo.invEXISTENCIALOTE WHERE IDBODEGA=@IdBodega AND IDPRODUCTO=@IdProducto AND IDLOTE=@IdLote))
	UPDATE dbo.invEXISTENCIALOTE SET EXISTENCIA=EXISTENCIA + @Cantidad 
	WHERE IDBODEGA=@IdBodega and IDPRODUCTO=@IdProducto and IDLOTE=@IdLote
ELSE	
	INSERT INTO dbo.invEXISTENCIALOTE(IDBODEGA,IDPRODUCTO,IDLOTE,EXISTENCIA)
	VALUES(@IdBodega,@IdProducto,@IdLote,@Cantidad)
GO 

CREATE PROCEDURE dbo.invUpdateExistenciaBodega(@IdBodega INT, @IdProducto INT,@Cantidad DECIMAL(28,8))
as 
IF (EXISTS(SELECT * FROM dbo.invEXISTENCIABODEGA WHERE IDBODEGA=@IdBodega AND IDPRODUCTO=@IdProducto))
	UPDATE dbo.invEXISTENCIALOTE SET EXISTENCIA=EXISTENCIA + @Cantidad 
	WHERE IDBODEGA=@IdBodega and IDPRODUCTO=@IdProducto
ELSE	
	INSERT INTO dbo.invEXISTENCIABODEGA(IDBODEGA,IDPRODUCTO,EXISTENCIA)
	VALUES(@IdBodega,@IdProducto,@Cantidad)


GO 

ALTER TABLE dbo.invCABMOVIMIENTOS ADD REFERENCIA NVARCHAR (20) NOT NULL

GO 

ALTER TABLE dbo.fafFACTURA_LINEA ADD IDLote INT  NOT NULL

GO 

--ALTER TABLE dbo.fafFACTURA_Linea DROP CONSTRAINT pkfafFACTURA_LINEA 
Alter Table dbo.fafFACTURA_LINEA add constraint pkfafFACTURA_LINEA primary key clustered ( IDBodega, IDFactura, IDCliente, 
		IDVendedor, Fecha, IDProducto,IDLote )

GO 

Alter Table dbo.fafFACTURA_LINEA add constraint fkfafFACTURA_LINEALote foreign key (IDLote)
references dbo.invLOTE (IDLote)

GO 
--drop procedure dbo.invUpdateExistenciaBodegaLinea
CREATE PROCEDURE [dbo].[invUpdateExistenciaBodegaLinea] @IdBodega INT, @IdProducto INT,@IdLote INT = NULL,@Cantidad DECIMAL(28,8), @IdTipoTransaccion INT,@Usuario nvarchar(50)
AS


DECLARE @TRANSACCION as NVARCHAR(20),@FACTOR AS SMALLINT
declare @i int,@iRwCnt int, @Lote int,  @CantidadLote decimal (28,8), @Completado bit, @CantidadAsignada decimal(28,8),@CantDemandada decimal(28,8),@CantidadDemandada AS DECIMAL(28,8)

SELECT @TRANSACCION= Transaccion,@FACTOR=Factor 
	FROM dbo.invTIPOMOVIMIENTO WHERE IDTipo=@IdTipoTransaccion

set @Cantidad = abs(@Cantidad) * @FACTOR

EXEC dbo.invUpdateExistenciaBodegaLote @IdBodega,@IdProducto,@IdLote,@Cantidad
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
INNER JOIN dbo.ccCLIENTE C ON fc.IDCliente=c.CodCliente
INNER JOIN dbo.invBODEGA B ON fc.IDBodega=b.IDBodega
WHERE FC.Anulada=0

GO 

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
SELECT IDBODEGA,IDPRODUCTO,IDLOTE,CANTIDAD,IDTIPO,COSTOLOCAL,COSTODOLAR
FROM dbo.vinvMovimientos  WHERE DOCUMENTO=@Documento AND IDPAQUETE=@IDPaquete

set @iRwCnt = @@ROWCOUNT
set @i = 1
set @Cantidad = 0 


while @i <= @iRwCnt 
	begin
		select @IDLote = IdLote, @Cantidad = Cantidad, @IdBodega= IdBodega, @IdProducto= IdProducto ,
				@IdTipoTransaccion=IDTipoTransaccion,@CostoLocal=CostoLocal,@CostoDolar = CostoDolar
		  from #tmpMovimiento where ID = @i
		exec dbo.invUpdateExistenciaBodegaLinea @IdBodega, @IdProducto,@IdProducto, @Cantidad,@IdTipoTransaccion,@Usuario
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

GO 


CREATE PROCEDURE dbo.invGetSugeridoLote(@IDBodega INT,@IDProducto INT,@Cantidad DECIMAL(28,8))	
AS 
/*SET @Cantidad =25
SET @IDProducto=1
SET @IDBodega=1
*/

SET NOCOUNT ON
declare @iRwCnt INT,@CantidadLote DECIMAL(28,8),@CantidadAsignada DECIMAL(28,8),@Completado BIT
DECLARE @i INT,@IDLote INT

Create Table #Resultado (
IDBodega nvarchar(20), --COLLATE Latin1_General_CI_AS, 
IDProducto nvarchar(20),-- COLLATE Latin1_General_CI_AS, 
IDLote int, Cantidad decimal(28,8) default 0 
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

SELECT A.IDBodega, A.IDProducto, A.IDLote,L.LoteInterno, L.LoteProveedor,
       L.FechaVencimiento, L.FechaFabricacion, A.Cantidad
  FROM #Resultado A
INNER JOIN dbo.invLOTE L ON L.IDLote = A.IDLote
	
DROP TABLE #Resultado
DROP TABLE #ProductoLote



