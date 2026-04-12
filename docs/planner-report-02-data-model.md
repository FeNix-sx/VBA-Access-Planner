# Planner: модель данных

Источник схем: JSON в [`DB_VBA/VBA_Tables`](../DB_VBA/VBA_Tables). Типы полей в JSON — числовые коды DAO (кратко: `1` Boolean, `3` Integer, `4` Long, `8` Date/Time, `10` Text).

Архитектура: **frontend** (приложение) + **backend** `Planner_BE.accdb` (путь задаётся в `tbTableConnections`, по умолчанию `CurrentProject.Path & "\BE\Planner_BE.accdb"` в [`modTableConnect.bas`](../DB_VBA/VBA_Modules/modTableConnect.bas)).

## Таблицы и смысл

### `tbEventInstances`

Фактические события календаря (одна строка — один экземпляр на дату).

| Поле | Назначение |
|------|------------|
| `InstanceID` | PK, AutoNumber |
| `EventDate` | Дата события |
| `EventNote` | Текст / описание |
| `Basis`, `BasisAttachment` | Основание и вложение |
| `CompletionDate`, `CompletionMark` | Выполнение |
| `LastModified` | Метка изменения |
| `AttachmentPath` | Путь к файлу |
| `ExecutorID` | FK на `tbExecutors` (индекс `relExecutorsEvents`) |
| `Notes` | Доп. заметки |

**Код:** загрузка в ячейки календаря — `Form_f_daily_planner` (`LoadEventData` и связанные процедуры); редактирование списка на день — `Form_frmDayEvents`.

### `tbExecutors`

Справочник исполнителей: `ID`, ФИО (`LastName`, `FirstName`, `MiddleName`), `Position`, `SortOrder`.

**Код:** комбобоксы на главной форме, в генераторе, в поиске, в форме дня; форма `frmExecutors`.

### `tbThemes`

Цветовые схемы: `ThemeID`, `ThemeName`, `IsActive`, набор полей `*_Text` / `*_Back` / `*_Border` для текущего/другого месяца, «сегодня», заголовка, `Form_Back`.

**Код:** `modDatabase` (создание/инициализация), `Form_f_daily_planner` (`LoadDefaultTheme`, `ApplyTheme`), `Theme_WritePalette` в `modThemeColors` для отчётов.

### `tbSettings`

Ключ-значение: `SettingName` (PK), `SettingValue` (текст).

Используемые ключи (по коду):

| SettingName | Где используется |
|-------------|------------------|
| `HideCompleted` | `f_daily_planner` — чекбокс скрытия выполненных |
| `ShowBirthdaysPanel` | `f_daily_planner` — показ панели ДР |
| `SelectedExecutor` | `f_daily_planner` — фильтр исполнителя |
| `ActivationKey`, `ActivationDate` | `modProtection` — привязка лицензии к ПК |
| `ComputerID`, `FirstRunDate`, `IsActivated`, `DaysUsed`, `LastRunDate` | `InitializeLicense` / тесты защиты |
| `AdminPassword` | хэш пароля администратора |

### `tbTableConnections`

Связь имён таблиц с путём к backend: `TableName`, `TablePath`, `Description`.

**Код:** `ConnectAllTables`, `AutoConnectOnStartup`, `MigrateAddBirthdaysConnectionIfMissing`.

### `tbTempEvents`

Черновики событий генератора до сохранения в `tbEventInstances`: `TempID`, `EventDate`, `EventNote`, `OriginalDay`, `ExecutorID`, вложения.

**Код:** `Form_frmEventGenerator` (`GenerateEvents`, `AddTempEvent`, `cmdSave_Click`).

### `tbPeriodicity`

Типы периодичности: `PeriodicityID`, `PeriodicityName`, `Description`.

**Код:** комбобокс периодичности в `frmEventGenerator`.

### `tbRules`

Правила генерации (справочник): `RuleID`, `RuleName`, `RuleDescription`, `CalculationMethod`.

**Код:** связан с логикой генератора (контекст в [`modDatabase.bas`](../DB_VBA/VBA_Modules/modDatabase.bas) в комментариях и процедурах создания таблиц).

### `tbBirthdays`

Дни рождения (в backend): `ID`, ФИО, `BirthDate`, `Notes`.

**Код:** `modBirthdays` (запрос `qryBirthdaysForPanel`, SQL `BirthdaysPanelRecordSourceSql`, отчёт панели), формы `frmBirthdaysList`, `frmBirthdayCard`.

Подробности развёртывания: [`Migrations/Migration_tbBirthdays.txt`](../DB_VBA/Migrations/Migration_tbBirthdays.txt).

## Упоминаемые в коде, но без JSON в этой выгрузке

В [`modDatabase.bas`](../DB_VBA/VBA_Modules/modDatabase.bas) в шапке перечислены **`tbEvents`** (устаревшая), **`tbRecurringEvents`** (шаблоны) — в каталоге `VBA_Tables` для них нет файлов; актуальная модель событий в репозитории сфокусирована на **`tbEventInstances`** и **`tbTempEvents`**.

## Связи (логические)

- `tbEventInstances.ExecutorID` → `tbExecutors.ID`
- Остальные таблицы связаны сценариями приложения, а не обязательно FK в Access (проверять в живой базе).

## Связанные отчёты

- [Обзор](planner-report-01-overview.md)
- [Формы и отчёты](planner-report-03-forms-and-reports.md)
- [Машинный индекс](planner-report-04-machine-index.yaml)
