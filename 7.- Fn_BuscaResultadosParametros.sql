Use db_Gen_des
Go

--
-- obetivo:     Consulta en la entidad dbo_parametros en la base de datos db_Gen_des
-- fecha:       15/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--

Create Or Alter Function dbo.Fn_BuscaResultadosParametros
   (@PnParametro        Integer,
    @PsColumna          Sysname)
Returns Varchar(500)

Begin
   Declare
      @o_Salida  Varchar ( 500)

   Begin
      If @PsColumna Not In ('valor', 'cadena')
         Begin
            Set @o_Salida = 'Error en Selección'
            Return (@o_Salida)
         End

      If @PsColumna = 'cadena'
         Begin
            Select @o_Salida = cadena
            From   dbo.dbo_Parametros With (Nolock)
            Where  id = @PnParametro
         End
      Else
         Begin
            Select @o_Salida = Cast(Valor As Varchar)
            From   dbo.dbo_Parametros With (Nolock)
            Where  id = @PnParametro
         End

      Return (@o_Salida)

   End

End
Go

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Función que Consulta la entidad dbo_parametros.',
   @w_procedimiento  NVarchar(250) = 'Fn_BuscaResultadosParametros';

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
