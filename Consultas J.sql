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
VALUES (2, 'FAC', 'Facturaci�n', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (3, 'AJE', 'Ajuste por Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (4, 'AJS', 'Ajuste por Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (5, 'BON', 'Bonificaci�n', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (6, 'PRS', 'Pr�stamo Salida', 'S', -1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (7, 'PRE', 'Pr�stamo Entrada', 'E', 1, 1)
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
VALUES (11, 'FIE', 'Ajuste F�sico Entrada', 'E', 1, 1)
GO
INSERT [dbo].[invTIPOMOVIMIENTO]( IDTipo, Transaccion , Descr, Naturaleza , Factor, ReadOnly )
VALUES (12, 'FIS', 'Ajuste F�sico Salida', 'S', -1, 1)
GO


--INSERTAR PAQUETES
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (1,'COM','Paquete de Compra',0,1,'COM000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (2,'FAC','Paquete de Facturaci�n',0,1,'FAC000000000000',1)
GO
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (3,'AJU','Paquete de Ajuste',0,1,'AJU000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (4,'BON','Paquete de Bonificaci�n',0,1,'BON000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (5,'PRE','Paquete de Pr�stamo',0,1,'PRE000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (6,'CON','Paquete de Consumo',0,1,'CON000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (7,'TRS','Paquete de Traslado',0,1,'TRS000000000000',1)
GO 
INSERT INTO dbo.invPAQUETE(IDPaquete, PAQUETE,  Descr , Consecutivo, ConsecAutomatico, Documento, Activo )
VALUES (8,'FIS','PPaquete de Ajuste F�sico',0,1,'FIS000000000000',1)
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


select P.IDPaquete, M.IDTipo, M.Transaccion 
from dbo.invPAQUETE P inner join dbo.invTIPOMOVIMIENTO M
on P.IDPaquete = M.IDTipo 

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



go 
create procedure dbo.UpdateExistenciaBodega(@IdBodega int, @IdProducto int, @Cantidad decimal(28,8),@IdTipoTransaccion as int)	
as
	declare @Factor int, @Transaccion nvarchar(5)
	
	select @Factor=Factor,@Transaccion=Transaccion 
	from dbo.invTIPOMOVIMIENTO where IDTipo=@IdTipoTransaccion
	
	
	update dbo.invEXISTENCIABODEGA set EXISTENCIA = EXISTENCIA + (@Cantidad * @Factor)
	where IDPRODUCTO=@IdProducto and IDBODEGA=@IdBodega


