Use Ejercicio_des
Go

--
-- obetivo:     Función que Calcula la Dirección IP.
-- fecha:       12/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--

Create Or Alter Function dbo.Fn_BuscaDireccionIP()
Returns Varchar(30)
With encryption 
As

Begin
   Declare
      @o_salida           Varchar(30)

   Begin
      Set @o_salida = Cast(ConnectionProperty('client_net_address') As Varchar(30))
   End

   Return(@o_salida)

End
Go

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Función que Calcula la Dirección IP.',
   @w_procedimiento  NVarchar(250) = 'Fn_BuscaDireccionIP';

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