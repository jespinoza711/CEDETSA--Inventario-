
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
CREATE PROCEDURE dbo.invInsertCabMovimientos  @IDPaquete int,
	@Documento nvarchar(20) OUTPUT,@Fecha DATETIME,@Concepto NVARCHAR(255),@UserInsert AS NVARCHAR(20),
	@UserUpdate  NVARCHAR(20)
	
as

set nocount ON

exec dbo.invGetNextConsecutivoPaquete @IDPaquete,@Documento OUTPUT



INSERT INTO dbo.invCABMOVIMIENTOS(IDPAQUETE, DOCUMENTO, FECHA, CONCEPTO,
            UserInsert, UserUpdate, FechaInsert, FechaUpdate)

VALUES (@IDPaquete,@Documento,@Fecha,@Concepto,@UserInsert,@UserUpdate,GETDATE(),GETDATE())


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

