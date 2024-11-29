Use Ejercicio_des
Go

/*
Declare
   @PnAnio                   Smallint     = 2023,
   @PnMes                    Tinyint      = 13,
   @PnEstatus                Integer      = 0,
   @PsMensaje                Varchar(250) = Null;
Begin
   Execute dbo.Spp_RecalculaSaldoAcumulados @PnAnio    = @PnAnio,
                                            @PnMes     = @PnMes,
                                            @PnEstatus = @PnEstatus  Output,
                                            @PsMensaje = @PsMensaje  Output;

   Select @PnEstatus IdError,  @PsMensaje MensajeError
   Return

End
Go

*/

--
-- obetivo:     Recalcular los saldos Acumulados a partir de un mes de inicio.
-- fecha:       12/11/2024.
-- Programador: Pedro Zambrano
-- Versión:     1
--

Create Or Alter Procedure dbo.Spp_RecalculaSaldoAcumulados
  (@PnAnio                   Smallint,
   @PnMes                    Tinyint,
   @PnEstatus                Integer      = 0    Output,
   @PsMensaje                Varchar(250) = Null Output)
As

Declare
   @w_tabla                   Sysname,
   @w_tablaAnt                Sysname,
   @w_anioIni                 Integer,
   @w_anioFin                 Integer,
   @w_anioAnte                Integer,
   @w_error                   Integer,
   @w_registros               Integer,
   @w_secuencia               Integer,
   @w_operacion               Integer,
   @w_mesIni                  Tinyint,
   @w_mesFin                  Tinyint,
   @w_mesAnte                 Integer,
   @w_mes                     Tinyint,
   @w_linea                   Integer,
   @w_comilla                 Char(1),
   @w_chmes                   Char(3),
   @w_desc_error              Varchar(  250),
   @w_mensaje                 Varchar(  Max),
   @w_idusuario               Varchar(  Max),
   @w_usuario                 Nvarchar(  20),
   @w_sql                     NVarchar(1500),
   @w_param                   NVarchar( 750),
   @w_ultactual               Datetime,
   @w_sucursable              Tinyint,
   @w_x                       Bit,
   @w_existeAux               Bit;

Begin
   Set Nocount            On
   Set Xact_Abort         On
   Set Ansi_Nulls         Off

   Select @PnEstatus         = 0,
          @PsMensaje         = Null,
          @w_operacion       = 9999,
          @w_anioIni         = @PnAnio,
          @w_mes             = @PnMes,
          @w_comilla         = Char(39),
          @w_x               = 0,
          @w_ultactual       = Getdate(),
          @w_existeAux       = 0,
          @w_sucursable      = Isnull(db_Gen_des.dbo.Fn_BuscaResultadosParametros(12, 'valor'), 0);

   Select @w_mesIni = Min(Valor),
          @w_mesFin = Max(Valor)
   From   dbo.catCriteriosTbl With (Nolock)
   Where  criterio  = 'mes';

   Select @w_anioFin = Max(Substring(name, 8, 4))
   from   sys.tables
   Where  Substring(name, 1, 7)  = 'control'
   And    Isnumeric(Substring(name, 8, 4))  = 1

--
-- Creación de Tablas Temporales.
--

   Create table #TempControl
  (secuencia    Integer Not Null Identity (1, 1) Primary Key,
   anio          Smallint       Not Null,
   mes           Tinyint        Not Null,
   tabla         Sysname        Not Null,
   tablaAnt      Sysname            Null);

--
-- Inicio Proceso
--

   Begin Transaction

       While @w_x  = 0
       Begin
          While @w_mes <= @w_mesFin
          Begin
             Select @w_registros   = 0,
                    @w_secuencia   = 0,
                    @w_chmes       = dbo.Fn_BuscaMes (@w_mes, 2),
                    @w_mesAnte     = @w_mes -1,
                    @w_anioAnte    = @w_anioIni,
                    @w_tabla       = Case When @w_mes = 0
                                          Then Concat('catIni', @w_anioIni)
                                          When @w_mes = 13
                                          Then Concat('catCie', @w_anioIni)
                                          Else Concat('cat', @w_chmes, @w_anioIni)
                                     End;

             If @w_mesAnte < @w_mesIni
                Begin
                   Select @w_mesAnte  = @w_mesFin,
                          @w_anioAnte = @w_anioAnte -1;
                End

             Select  @w_chmes    = dbo.Fn_BuscaMes (@w_mesAnte, 2),
                     @w_tablaAnt = Case When @w_mesAnte = 0
                                        Then Concat('catIni', @w_anioAnte)
                                        When @w_mesAnte = 13
                                        Then Concat('catCie', @w_anioAnte)
                                        Else Concat('cat', @w_chmes, @w_anioAnte)
                                   End;

             If Ejercicio_DES.dbo.Fn_existe_tabla( @w_tabla ) = 0
                Begin
                   Goto Siguiente
                End

             Insert Into #TempControl
            (anio, mes, tabla, tablaAnt)
             Select @w_anioIni, @w_mes, @w_tabla, @w_tablaAnt;
             Set @w_registros = @@Identity

Siguiente:

             Set @w_mes = @w_mes + 1;

             If @w_mes > @w_mesFin
                Begin
                   Select @w_mes     = @w_mesIni,
                          @w_anioIni = @w_anioIni + 1;

                   If @w_anioIni > @w_anioFin
                      Begin
                         Set @w_x = 1
                         Break
                      End

                End

          End

       End

       While @w_secuencia < @w_registros
       Begin
          Select @w_secuencia += 1;

          Select @w_tabla    = tabla,
                 @w_tablaAnt = tablaAnt
          From   #TempControl
          Where  secuencia = @w_secuencia;
          If @@Rowcount = 0
             Begin
                Break
             End

         If @w_secuencia = 1
            Begin
               Set @w_sql = Concat('Update a ',
                                   'Set    Sact = SAnt + car - Abo ',
                                   'From   dbo.', @w_tabla, ' a With (Nolock) ')

               Begin Try
                  Execute (@w_sql)
               End   Try

               Begin Catch
                  Select  @w_Error      = @@Error,
                          @w_linea      = Error_line(),
                          @w_desc_error = Substring (Error_Message(), 1, 200)

               End Catch

               If IsNull(@w_error, 0) <> 0
                  Begin
                     Rollback Transaction
                     Select @PnEstatus = @w_error,
                            @PsMensaje = Concat('Error.: ', @w_Error, ' ', @w_desc_error, ' en Línea ', @w_linea);

                     Set Xact_Abort Off
                     Goto Salida
                  End

               If @w_sucursable = 2
                  Begin

--
-- Actualiza CatAux
--

                     Set @w_tabla = Replace(@w_tabla, 'cat', 'catAux');
                     If  Ejercicio_DES.dbo.Fn_existe_tabla(@w_tabla) = 0
                         Begin
                            Goto Proximo
                         End

                     Set @w_sql = Concat('Update a ',
                                         'Set    Sact = SAnt + car - Abo ',
                                         'From   dbo.', @w_tabla, ' a With (Nolock) ')
                     Begin Try
                        Execute (@w_sql)
                     End   Try

                     Begin Catch
                        Select  @w_Error      = @@Error,
                                @w_linea      = Error_line(),
                                @w_desc_error = Substring (Error_Message(), 1, 200)

                     End Catch

                     If IsNull(@w_error, 0) <> 0
                        Begin
                           Rollback Transaction
                           Select @PnEstatus = @w_error,
                                  @PsMensaje = Concat('Error.: ', @w_Error, ' ', @w_desc_error, ' en Línea ', @w_linea);

                           Set Xact_Abort Off
                           Goto Salida
                        End

                  End
            End

         Else
            Begin
               Set @w_sql = Concat('Update dbo.', @w_tabla, ' ',
                                   'Set    Sant = b.Sact,     ',
                                          'Sact = b.Sact + a.car - a.abo  ',
                                   'From   dbo.', @w_tabla,    ' a With (Nolock) ',
                                   'Join   dbo.', @w_tablaAnt, ' b With (Nolock) ',
                                   'On     b.llave  = a.llave ',
                                   'And    b.moneda = a.moneda ',
                                   'And    b.Niv    = a.Niv');
               Begin Try
                  Execute (@w_sql)
               End Try

               Begin Catch
                  Select  @w_Error      = @@Error,
                          @w_linea      = Error_line(),
                          @w_desc_error = Substring (Error_Message(), 1, 200)

               End Catch

               If IsNull(@w_error, 0) <> 0
                  Begin
                     Rollback Transaction
                     Select @PnEstatus = @w_error,
                            @PsMensaje = Concat('Error.: ', @w_Error, ' ', @w_desc_error, ' en Línea ', @w_linea);

                     Set Xact_Abort Off
                     Goto Salida
                  End

--

               If @w_sucursable = 2
                  Begin

--
-- Actualiza CatAux
--

                     Select @w_tablaAnt = Replace(@w_tablaAnt, 'cat', 'catAux'),
                            @w_tabla    = Replace(@w_tabla,    'cat', 'catAux');

                     If  Ejercicio_DES.dbo.Fn_existe_tabla(@w_tabla) = 0
                         Begin
                            Goto Proximo
                         End
                         
                     If  Ejercicio_DES.dbo.Fn_existe_tabla(@w_tablaAnt) = 0
                         Begin
                            Goto Proximo
                         End

                   Set @w_sql = Concat('Update dbo.', @w_tabla, ' ',
                                       'Set    Sant = b.Sact, ',
                                              'Sact = b.Sact + a.car - a.abo ',
                                       'From   dbo.', @w_tabla,    ' a With (Nolock) ',
                                       'Join   dbo.', @w_tablaAnt, ' b With (Nolock) ',
                                       'On     b.llave       = a.llave ',
                                       'And    b.moneda      = a.moneda ',
                                       'And    b.Niv         = a.Niv ',
                                       'And    b.Sector_id   = a.Sector_id ',
                                       'And    b.Sucursal_id = a.Sucursal_id');

                   Begin Try
                      Execute (@w_sql)
                   End Try

                   Begin Catch
                      Select  @w_Error      = @@Error,
                              @w_linea      = Error_line(),
                              @w_desc_error = Substring (Error_Message(), 1, 200)

                   End Catch

                   If IsNull(@w_error, 0) <> 0
                      Begin
                         Rollback Transaction
                         Select @PnEstatus = @w_error,
                                @PsMensaje = Concat('Error.: ', @w_Error, ' ', @w_desc_error, ' en Línea ', @w_linea);

                         Set Xact_Abort Off
                         Goto Salida
                      End
                End

            End

Proximo:

       End

   Commit Transaction;

Salida:

   Set Xact_Abort  On
   Return

End
Go

--
-- Comentarios
--

Declare
   @w_valor          Nvarchar(250) = 'Procedimiento que Recalcula los saldos contables acumulados.',
   @w_procedimiento  NVarchar(250) = 'Spp_RecalculaSaldoAcumulados';

If Not Exists (Select Top 1 1
               From   sys.extended_properties a
               Join   sysobjects  b
               On     b.xtype   = 'P'
               And    b.name    = @w_procedimiento
               And    b.id      = a.major_id)
   Begin
      Execute  sp_addextendedproperty @name       = N'MS_Description',
                                      @value      = @w_valor,
                                      @level0type = 'Schema',
                                      @level0name = N'dbo',
                                      @level1type = 'Procedure',
                                      @level1name = @w_procedimiento

   End
Else
   Begin
      Execute sp_updateextendedproperty @name       = 'MS_Description',
                                        @value      = @w_valor,
                                        @level0type = 'Schema',
                                        @level0name = N'dbo',
                                        @level1type = 'Procedure',
                                        @level1name = @w_procedimiento
   End
Go

