Use Ejercicio_des
Go

--
-- obetivo:      Consulta en la entidad catCriteriosTbl las iniciales del mes
-- fecha:       12/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--


Create  Or Alter Function dbo.Fn_BuscaMes
 (@PnMes          Smallint,
  @PnTipoConsulta Tinyint)
Returns Varchar(15)

As
Begin
    Declare
       @w_salida   Varchar(15)

       Select @w_salida = Case When @PnTipoConsulta = 1
                               Then descripcion
                               Else valorAdicional
                          End
       From   dbo.catCriteriosTbl
       Where  criterio = 'mes'
       And    valor    = @PnMes
       If @@Rowcount = 0
          Begin
             Set @w_salida = Char(32);
          End

Return @w_salida

End
Go

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Función que Consulta en la entidad catCriteriosTbl las iniciales del mes.',
   @w_procedimiento  NVarchar(250) = 'Fn_BuscaMes';

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


