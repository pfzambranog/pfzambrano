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
   @w_tablaAux                Sysname,
   @w_anioIni                 Integer,
   @w_anioFin                 Integer,
   @w_error                   Integer,
   @w_registros               Integer,
   @w_operacion               Integer,
   @w_mesIni                  Tinyint,
   @w_mesFin                  Tinyint,
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

   Create Table #TempSaldos
  (anio          Smallint       Not Null,
   mes           Tinyint        Not Null,
   tabla         Sysname        Not Null,
   llave         Varchar(20)    Not Null,
   moneda        Varchar( 2)    Not Null,
   Sant          Decimal(18, 2) Not Null,
   car           Decimal(18, 2) Not Null,
   Abo           Decimal(18, 2) Not Null,
   Sact          Decimal(18, 2) Not Null Default  0,
   Primary Key (anio, mes, tabla, llave, moneda));

   Create Table #TempSaldosAux
  (anio          Smallint       Not Null,
   mes           Tinyint        Not Null,
   tabla         Sysname        Not Null,
   llave         Varchar(20)    Not Null,
   moneda        Varchar( 2)    Not Null,
   Sector_id     Nvarchar(4)    Not Null,
   Sucursal_id   Integer        Not Null,
   Sant          Decimal(18, 2) Not Null,
   car           Decimal(18, 2) Not Null,
   Abo           Decimal(18, 2) Not Null,
   Sact          Decimal(18, 2) Not Null Default  0,
   Primary Key (anio, mes, tabla, llave, moneda, Sector_id, Sucursal_id));

--
-- Inicio Proceso
--

   Begin Transaction

       While @w_x  = 0
       Begin
          While @w_mes <= @w_mesFin
          Begin
             Select @w_registros   = 0,
                    @w_chmes       = dbo.Fn_BuscaMes (@w_mes, 2),
                    @w_tabla       = Case When @w_mes = 0
                                          Then Concat('catIni', @w_anioIni)
                                          When @w_mes = 13
                                          Then Concat('catFin', @w_anioIni)
                                          Else Concat('cat', @w_chmes, @w_anioIni)
                                     End,
                    @w_tablaAux    = Case When @w_mes = 0
                                          Then Concat('catAuxIni', @w_anioIni)
                                          When @w_mes = 13
                                          Then Concat('catAuxFin', @w_anioIni)
                                          Else Concat('catAux', @w_chmes, @w_anioIni)
                                     End;

             If Ejercicio_DES.dbo.Fn_existe_tabla( @w_tabla ) = 0
                Begin
                   Goto Siguiente
                End


             If @w_sucursable != 2
                Begin
                   Set @w_existeAux = 0;
                End
             Else
                Begin
                   Set @w_existeAux = Ejercicio_DES.dbo.Fn_existe_tabla(@w_tablaAux);
                End;

--
-- Actualiza saldo inicial cat
--
             Set @w_sql = Concat('Update dbo.', @w_tabla, ' ',
                                 'Set    Sant = b.Sact ',
                                 'From   dbo.', @w_tabla, ' a With (Nolock) ',
                                 'Join   #TempSaldos        b ',
                                 'On     b.llave  = a.llave ',
                                 'And    b.moneda = a.moneda');

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

--
-- Actualiza saldo inicial catAux
--

             If @w_existeAux = 1
                Begin
                   Set @w_sql = Concat('Update dbo.', @w_tabla, ' ',
                                       'Set    Sant = b.Sact ',
                                       'From   dbo.', @w_tablaAux, ' a With (Nolock) ',
                                       'Join   #TempSaldosAux        b ',
                                       'On     b.llave       = a.llave ',
                                       'And    b.moneda      = a.moneda ',
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

--
-- Inicializa temporal Cat
--

             Set @w_sql = 'Truncate Table #TempSaldos'

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
-- Inicializa temporal CatAux
--

             Set @w_sql = 'Truncate Table #TempSaldosAux'

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
-- Inserta datos de cat
--

             Set @w_sql = Concat('Select ', @w_anioIni, ', ',  @w_mes,        ', ',
                                            @w_comilla, @w_tabla, @w_comilla, ', ',
                                           'llave, moneda , Sant, car, Abo ',
                                 'From   dbo.', @w_tabla, ' With (Nolock)');

             Begin Try
                Insert Into #TempSaldos
                (anio,    mes,  tabla, llave,
                 moneda,  Sant, car,   Abo)
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

             Update #TempSaldos
             Set    Sact = Sant + car - Abo


--
-- Inserta datos de catAux
--

             If @w_existeAux = 1
                Begin
                   Set @w_sql = Concat('Select ', @w_anioIni, ', ',  @w_mes,        ', ',
                                                  @w_comilla, @w_tabla, @w_comilla, ', ',
                                                 'llave, moneda , Sector_id, Sucursal_id, Sant, car, Abo ',
                                       'From   dbo.', @w_tablaAux, ' With (Nolock)');

                   Begin Try
                      Insert Into #TempSaldosAux
                      (anio,    mes,       tabla,       llave,
                       moneda,  Sector_id, Sucursal_id, Sant,
                       car,   Abo)
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

             Update #TempSaldosAux
             Set    Sact = Sant + car - Abo

--
--
--
             If @w_mes = @PnMes
                Begin
--
-- Actualiza Cat
--
                   Set @w_sql = Concat('Update a ',
                                       'Set    Sact = b.Sact ',
                                       'From   dbo.', @w_tabla, ' a With (Nolock) ',
                                       'Join   #TempSaldos b ',
                                       'On     b.llave  = a.llave ',
                                       'And    b.moneda = a.moneda')
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

                   If @w_existeAux = 1
                      Begin
--
-- Actualiza CatAux
--
                          Set @w_sql = Concat('Update a ',
                                              'Set    Sact = b.Sact ',
                                              'From   dbo.', @w_tablaAux, ' a With (Nolock) ',
                                              'Join   #TempSaldosAux b ',
                                              'On     b.llave       = a.llave ',
                                              'And    b.moneda      = a.moneda ',
                                              'And    b.Sector_id   = a.Sector_id ',
                                              'And    b.Sucursal_id = a.Sucursal_id')
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
--
-- Actualiza Cat
--
                   Set @w_sql = Concat('Update a ',
                                       'Set    Sant = b.Sant, ',
                                              'Sact = b.sant + b.car - b.abo ',
                                       'From   dbo.', @w_tabla, ' a With (Nolock) ',
                                       'Join   #TempSaldos b ',
                                       'On     b.llave  = a.llave ',
                                       'And    b.moneda = a.moneda')
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

--
-- Actualiza CatAux
--
                   If @w_existeAux = 1
                      Begin
                         Set @w_sql = Concat('Update a ',
                                             'Set    Sant = b.Sant, ',
                                                    'Sact = b.sant + b.car - b.abo ',
                                             'From   dbo.', @w_tablaAux, ' a With (Nolock) ',
                                             'Join   #TempSaldosAux b ',
                                             'On     b.llave       = a.llave ',
                                             'And    b.moneda      = a.moneda ',
                                             'And    b.Sector_id   = a.Sector_id ',
                                             'And    b.Sucursal_id = a.Sucursal_id')
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

