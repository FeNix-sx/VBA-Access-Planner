# Planner: формы, отчёты, сценарии UI

Код форм: [`DB_VBA/VBA_FormCode`](../DB_VBA/VBA_FormCode). Контролы: [`DB_VBA/VBA_FormControls`](../DB_VBA/VBA_FormControls). Отчёт: [`DB_VBA/VBA_ReportCode`](../DB_VBA/VBA_ReportCode), [`DB_VBA/VBA_ReportControls`](../DB_VBA/VBA_ReportControls).

## Главная форма `f_daily_planner`

Файл: [`Form_f_daily_planner.bas`](../DB_VBA/VBA_FormCode/Form_f_daily_planner.bas).

**Загрузка:** лицензия, автоподключение BE, тема, настройки, построение календаря, раскладка панели дней рождения.

**Навигация:** `btn_current_Click`, `btn_next_Click`, `btn_previous_Click`; публичные `GoToNextMonth`, `GoToPreviousMonth`, `GoToCurrentMonth` (для демо).

**Открытие других форм:**

| Элемент / действие | Цель |
|--------------------|------|
| `cmdEvengGenerate_Click` | `frmEventGenerator` |
| `cmdExecutors_Click` | `frmExecutors` |
| `btn_BirthdaysManage_Click` | `frmBirthdaysList` |
| `cmdRunDemo_Click` | `frmDemo` |
| `btn_theme_Click` | `frmThemeSelector` (modal) |
| `cmdSearchEvents_Click` | `frmSearch` |
| DblClick по полям дней `fld_day_*` | `frmDayEvents` для выбранной даты |

**Публичные методы для внешних вызовов:** `BuildCalendar`, `ApplyTheme`, `InitializeExecutorFilter`, `OpenDayEvents`, `ApplyExecutorFilter`, `ApplyHideCompletedFilter`, `GoToNextMonth`, `GoToPreviousMonth`, `GoToCurrentMonth`.

**Тема:** `ApplyTheme` читает строку из `tbThemes`, красит контролы и вызывает `Theme_WritePalette` (`modThemeColors`).

## `frmDayEvents` — события выбранного дня

Файл: [`Form_frmDayEvents.bas`](../DB_VBA/VBA_FormCode/Form_frmDayEvents.bas).

Режимы просмотр/редактирование, переход на предыдущий/следующий день, отметки выполнения, вложения через **`frmFileFolderSelector`** (диалог). Публичные методы для автоматизации: `GoToNextDay`, `GoToPreviousDay`, `StartEditMode`, `SaveChanges`, `CloseForm`.

**Откуда открывается:** `f_daily_planner`, `frmSearch` (двойной клик по результату → `OpenDayFromSearch`).

## `frmSearch` — поиск событий

Файл: [`Form_frmSearch.bas`](../DB_VBA/VBA_FormCode/Form_frmSearch.bas).

Фильтры по дате, тексту, исполнителю, статусу выполнения, вложениям; `BuildSearchConditions`; сброс. Публичные: `ExecuteSearch`, `ResetSearch`, `CloseSearchForm` (для демо).

## `frmEventGenerator` — генератор повторов

Файл: [`Form_frmEventGenerator.bas`](../DB_VBA/VBA_FormCode/Form_frmEventGenerator.bas).

Выбор периодичности, диапазона дат, исполнителя, дня месяца/недели; генерация в **`tbTempEvents`**; сохранение в **`tbEventInstances`** и обновление календаря (`UpdateCalendarForm`). Вложения — через `frmFileFolderSelector`.

## `frmExecutors`

Файл: [`Form_frmExecutors.bas`](../DB_VBA/VBA_FormCode/Form_frmExecutors.bas).

Ленточная форма над `tbExecutors`; при закрытии обновляет фильтр на главной (`InitializeExecutorFilter` в `f_daily_planner`).

## `frmThemeSelector`

Файл: [`Form_frmThemeSelector.bas`](../DB_VBA/VBA_FormCode/Form_frmThemeSelector.bas).

Список тем из `tbThemes`; `btnApply` вызывает **`Forms!f_daily_planner.ApplyTheme`** и закрывает форму. Горячая клавиша открывает **`frmPassword`** (админ-поток). Публичные методы для демо: `ApplySelectedTheme`, `CloseThemeForm`, `GetThemeCount`, `SelectThemeByIndex`.

## `frmPassword` / `frmAdmin`

- [`Form_frmPassword.bas`](../DB_VBA/VBA_FormCode/Form_frmPassword.bas) — ввод пароля; при успехе открывает `frmAdmin`.
- [`Form_frmAdmin.bas`](../DB_VBA/VBA_FormCode/Form_frmAdmin.bas) — краткий админ-интерфейс (кнопки `cmdAdmin` / `cmdOffAdmin` — детали в файле).

## Дни рождения: `frmBirthdaysList`, `frmBirthdayCard`

- **Список:** [`Form_frmBirthdaysList.bas`](../DB_VBA/VBA_FormCode/Form_frmBirthdaysList.bas) — добавить (новая карточка), редактировать, удалить, DblClick по полям открывает карточку.
- **Карточка:** [`Form_frmBirthdayCard.bas`](../DB_VBA/VBA_FormCode/Form_frmBirthdayCard.bas) — запись в `tbBirthdays`, сохранение/закрытие.

После правок список может вызывать `RefreshBirthdaysUIAfterEdit` из `modBirthdays`.

## `frmDemo`

Файл: [`Form_frmDemo.bas`](../DB_VBA/VBA_FormCode/Form_frmDemo.bas).

Пошаговый таймерный тур по функциям (навигация, события, фильтры, поиск, темы). Много `Demo_*` процедур; открывает `f_daily_planner`, `frmSearch`, `frmThemeSelector` и имитирует действия пользователя.

## `frmFileFolderSelector`

Модальный выбор файла или папки; возвращает путь вызывающей форме (`frmDayEvents`, `frmEventGenerator`). Открытие с `OpenArgs` `"Main"` / `"Basis"`.

## Отчёт `rptBirthdays`

Файл: [`Report_rptBirthdays.bas`](../DB_VBA/VBA_ReportCode/Report_rptBirthdays.bas).

События `Report_Open`, `Report_Activate`, секции `GroupHeader`, `Detail` — раскраска по теме из **`modThemeColors`** (палитра активной темы), группировка по близости даты ДР. Источник записей в рантайме должен соответствовать запросу из `modBirthdays` (`BirthdaysPanelRecordSourceSql` / `EnsureQryBirthdaysForPanel`).

Доп. текст разметки: `Report_rptBirthdays.txt` (экспорт Access).

## Класс `cLogger`

Файл: [`cLogger.cls`](../DB_VBA/VBA_Classes/cLogger.cls).

Уровни логов, запись в файл (`WriteLog`, `DebugLog`, `InfoLog`, …). Может использоваться точечно в проекте (поиск по ссылкам `New cLogger` в импортированной базе).

## Объект `Form2`

В выгрузке есть [`Form2_controls.json`](../DB_VBA/VBA_FormControls/Form2_controls.json) без соответствующего `Form_Form2.bas` — вероятно заготовка или переименованная форма; в отчётах по функционалу не опираться без проверки в `.accdb`.

## Связанные отчёты

- [Обзор](planner-report-01-overview.md)
- [Модель данных](planner-report-02-data-model.md)
- [Машинный индекс](planner-report-04-machine-index.yaml)
