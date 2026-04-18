# Аудит VBA кода: формы и модули

Дата: 2026-04-16  
Область: `DB_VBA/VBA_FormCode/*.bas`, `DB_VBA/VBA_Modules/*.bas`  
Приоритет: синхронизация БД, качество кода, дублирование

## Что проверено

- Логика синхронизации/подключения: `modSchemaSync.bas`, `modTableConnect.bas`, связанный код в `modDatabase.bas`.
- Крупные модули/формы (в т.ч. >1000 строк).
- Потенциальные дубли, мертвый код, копипаст-расхождения.

## Подтвержденные находки (без изменения функционала)

### HIGH

- [x] **Копипаст-ошибка в `modProtection.bas`**  
      В `DisableShiftKey()` при `PROPERTY_NOT_FOUND` свойство создается как `AllowByPassKey = True`, хотя логика процедуры требует `False`.  
      Риск: обход Shift может остаться включенным после «отключения».

- [x] **Дублирующаяся функция проверки таблицы**  
      `DbHasTable` реализована и в `modSchemaSync.bas`, и в `modDatabase.bas` разными способами.  
      Риск: разный edge-case behavior при миграциях/инициализации.

### MEDIUM

- [x] **Повтор списка управляемых таблиц в `modTableConnect.bas`**  
      Одни и те же имена таблиц (`tbEventInstances`, `tbExecutors`, `tbTempEvents`, `tbPeriodicity`, `tbRules`, `tbBirthdays`) повторяются в нескольких SQL и процедурах.  
      Риск: «тихий» рассинхрон при добавлении/удалении таблиц.

- [x] **Неиспользуемая процедура `EnsureTable` в `modSchemaSync.bas`**  
      По коду вызовов нет; фактическая логика идет через `EnsureField`.  
      Риск: низкий, но создает ложную точку входа.

- [ ] **Дубли в `Form_frmDayEvents.bas`**  
      В `Form_Load` и `Form_Open` повторяется установка размеров/режима, причем с разными размерами (`18500x11000` vs `19500x9000`).  
      Риск: нестабильное поведение UI при открытии формы.

- [ ] **Дубли SQL-базы в `Form_frmSearch.bas`**  
      Базовый SELECT повторяется минимум в нескольких обработчиках (`Form_Load`, `cmdReset_Click`, `cmdSearch_Click`).  
      Риск: расхождение результата поиска после частичных правок.

### LOW

- [ ] **Сверхбольшие файлы, осложняющие поддержку**  
      - `Form_f_daily_planner.bas` (~2264)  
      - `Form_frmDemo.bas` (~1768)  
      - `modProjectAnalysis.bas` (~1489)  
      - `modDatabase.bas` (~1217)  
      - `mod_ExportVBACode.bas` (~1148)  
      Риск: высокая цена изменений и регрессий.

- [ ] **Серии из 84 обработчиков дней в `Form_f_daily_planner.bas`**  
      Это частично обусловлено Access-событиями, но требует строгого шаблонного сопровождения.

## Нюансы синхронизации (приоритетно)

- [x] Привести к единому источнику список таблиц для подключения/нормализации (`GetManagedTables()`).
- [x] Унифицировать `DbHasTable` (один канонический helper, остальное — обертки).
- [x] Проверить write-операции в подключении/миграции на использование `dbFailOnError`.
- [ ] Для пакетных шагов синхронизации рассмотреть транзакцию (там, где реально несколько взаимосвязанных изменений должны быть атомарны).
- [ ] Согласовать единый контракт ошибок: где показываем `MsgBox`, где только логируем и пробрасываем.

## Рекомендуемый безопасный план (без изменения поведения)

### Этап 1 — точечные фиксы

- [x] Исправить `DisableShiftKey()` для `PROPERTY_NOT_FOUND`: создавать `AllowByPassKey=False`.
- [x] Добавить комментарий `Deprecated` к `EnsureTable` (или удалить в отдельном шаге после повторной проверки использования).

### Этап 2 — устранение дублей синхронизации

- [x] Вынести `GetManagedTables()` и использовать его в:
      - `EnsureRequiredTableConnections`
      - выборке в `ConnectAllTables`
      - начальном заполнении таблицы подключений.
- [x] Вынести общий `DbHasTable` helper и свести два варианта к одному.

### Этап 3 — улучшение качества крупных файлов

- [ ] `Form_f_daily_planner.bas`: оставить как UI-фасад, вынести настройки/календарную логику/диагностику по модулям.
- [ ] `modDatabase.bas`: разделить DDL/seed/migration/backend-resolution.
- [ ] `mod_ExportVBACode.bas`: разделить coordinator/io/export-handlers.

## Минимальный smoke-check после правок

- [ ] `ConnectAllTables` на «чистом» FE и FE с уже заполненной `tbTableConnections`.
- [ ] Сценарий «backend путь не найден» + повторный выбор файла.
- [ ] Сценарий `SyncDatabaseSchema` для отсутствующих таблиц/полей.
- [ ] Открытие форм: `f_daily_planner`, `frmDayEvents`, `frmSearch`.
- [ ] Проверка bypass Shift: enable/disable и фактическое поведение при старте.

---

## Статус выполнения на 2026-04-17

Сделано в рамках безопасного Этапа 1–2 (без расширения scope):

- `fix`: `DisableShiftKey` при `PROPERTY_NOT_FOUND` создает `AllowByPassKey=False`.
- `refactor`: унификация `DbHasTable` через канонический helper (`modSchemaSync`) + wrapper в `modDatabase`.
- `refactor`: единый список управляемых таблиц через `GetManagedTables()` и переиспользование в ключевых точках `modTableConnect`.
- `fix`: добавлен `dbFailOnError` в write-операции по `tbTableConnections` в `modTableConnect`.
- `chore`: `EnsureTable` помечена как `Deprecated` в `modSchemaSync`.

Коммиты по шагам:

- `d4cdd1b` — fix: корректно отключать bypass Shift при отсутствии свойства
- `0c072bd` — refactor: унифицировать DbHasTable через канонический helper
- `a7f1e62` — refactor: вынести единый список управляемых таблиц
- `f0fec08` — fix: включить dbFailOnError для write-операций подключений
- `bd84381` — chore: пометить EnsureTable как deprecated

Итог: изменения **требуются** (в первую очередь по `DisableShiftKey` и консолидации синхронизации).  
Если нужно, следующим шагом можно применить только безопасные пункты Этапа 1–2 отдельными маленькими коммитами.
