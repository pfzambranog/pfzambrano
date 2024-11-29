Use Ejercicio_des
Go

--
-- obetivo:     Función que Calcula la Dirección Mac.
-- fecha:       12/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--

Create Or Alter Function dbo.Fn_Busca_DireccionMac()
Returns Char(30)
As

Begin
   Declare
      @o_salida           Varchar(30)

   Begin
      Select @o_salida = net_address
      From   master.dbo.sysprocesses 
      Where  spid = @@spid

      Set @o_salida = Isnull(@o_salida, ' ')
   End

   Return(@o_salida)

End
Go 

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Función que Calcula la DIrección MAC.',
   @w_procedimiento  NVarchar(250) = 'Fn_Busca_DireccionMac';

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