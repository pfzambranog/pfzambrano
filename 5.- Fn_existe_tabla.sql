Use Ejercicio_des
Go

--
-- obetivo:     que Valida la existencia de una tabla en la BD
-- fecha:       12/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--

Create Or Alter Function dbo.Fn_existe_tabla  
(@PsTabla      Sysname)  
Returns Smallint  
As  
  
Begin  
   Declare  
      @w_existe      Tinyint  = 0  
  
   Begin  
      If Exists ( Select Top 1 1  
                  From   Sysobjects  
                  Where  Uid  = 1  
                  And    Type = 'U'  
                  And    Name = @PsTabla)  
         Begin  
            Set @w_existe = 1  
         End  
  
   End  
  
   Return(@w_existe)  
  
End  
Go 

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Función que Valida la existencia de una tabla en la BD.',
   @w_procedimiento  NVarchar(250) = 'Fn_existe_tabla';

If Not Exists (Select Top 1 1
               From   sys.extended_properties a
               Join   sysobjects  b
               On     b.xtype   = 'Fn'
               And    b.name    = @w_procedimiento
               And    b.id      = a.major_id)
   Begin
      Execute  sp_addextendedproperty @name       = N'MS_Description',
                                      @value      = @w_valor,
                                      @level0type = 'Schema',
                                      @level0name = N'dbo',
                                      @level1type = 'Function', 
                                      @level1name = @w_procedimiento

   End
Else
   Begin
      Execute sp_updateextendedproperty @name       = 'MS_Description',
                                        @value      = @w_valor,
                                        @level0type = 'Schema',
                                        @level0name = N'dbo',
                                        @level1type = 'Function', 
                                        @level1name = @w_procedimiento
   End
Go 