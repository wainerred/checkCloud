# Selectel Server Inventory Script

Скрипт получает список всех серверов из аккаунта Selectel через API
и сохраняет результаты в **JSON** и **Excel (.xlsx)**.

---

## Что собирается

| Поле             | Выделенные серверы | Облачные серверы |
|------------------|--------------------|------------------|
| Имя сервера      | ✅                  | ✅                |
| IP-адрес         | ✅                  | ✅ (floating IP)  |
| Конфигурация     | ✅ CPU/RAM/Storage  | ✅ flavor/vCPU/RAM|
| Регион           | ✅ из location API  | ✅ availability zone|
| Операционная система | ✅             | ✅                |
| Дата оплаты      | ✅ paid_till        | — (почасовой биллинг)|

---

## Установка зависимостей

```bash
pip install requests openpyxl
```

---

## Получение API-токена (IAM Token)

1. Войдите в [my.selectel.ru](https://my.selectel.ru)
2. Перейдите: **Профиль → API → Создать токен**
3. Скопируйте токен

---

## Запуск

### Вариант 1 — через переменную окружения (рекомендуется)

```bash
export SELECTEL_TOKEN="ваш_iam_токен_здесь"
python selectel_servers.py
```

### Вариант 2 — через аргумент командной строки

```bash
python selectel_servers.py --token ваш_iam_токен_здесь
```

### Вариант 3 — с Keystone-токеном для облачных серверов

Если автоматическое получение Keystone-токена не работает
(бывает при некоторых настройках аккаунта):

```bash
python selectel_servers.py \
  --token ВАШ_IAM_ТОКЕН \
  --keystone-token ВАШ_KEYSTONE_ТОКЕН
```

Keystone-токен получается через панель управления Selectel:
**Мой аккаунт → API → OpenStack → Скачать openrc** или через запрос:

```bash
curl -s -X POST https://api.selectel.ru/identity/v3/auth/tokens \
  -H "Content-Type: application/json" \
  -d '{"auth":{"identity":{"methods":["token"],"token":{"id":"ВАШ_IAM_ТОКЕН"}}}}' \
  -i | grep X-Subject-Token
```

---

## Результат

Скрипт создаёт два файла с временной меткой:

```
selectel_servers_20250428_143022.json
selectel_servers_20250428_143022.xlsx
```

### Структура Excel

- **Лист 1** — "Выделенные серверы"
- **Лист 2** — "Облачные серверы"

Каждый лист содержит отформатированную таблицу со столбцами:
`Имя сервера | IP-адрес | Конфигурация | Регион | ОС | Дата оплаты`

### Структура JSON

```json
{
  "generated_at": "2025-04-28T14:30:22Z",
  "dedicated_servers": [
    {
      "name": "server-01",
      "ip": "185.1.2.3",
      "configuration": "CPU: Intel Xeon E5-2680 v4 | RAM: 128GB | Storage: 2x960GB SSD",
      "region": "Санкт-Петербург",
      "os": "Ubuntu 22.04 LTS",
      "billing_date": "2025-05-15",
      "uuid": "...",
      "status": "active"
    }
  ],
  "cloud_servers": [
    {
      "name": "web-server-prod",
      "ip": "95.2.3.4",
      "configuration": "SL1.2-2 (2 vCPU, 2GB RAM, 20GB disk)",
      "region": "ru-1a",
      "os": "Ubuntu 22.04 LTS",
      "billing_date": "—",
      "uuid": "...",
      "status": "ACTIVE"
    }
  ]
}
```

---

## Примечания

- **Облачные серверы** оплачиваются **почасово**, поэтому поле `billing_date`
  отображается как `—`. Данные о расходах смотрите в разделе **Биллинг**
  панели управления или через [Billing API](https://docs.selectel.ru/balance-and-payments/).

- Если в аккаунте много проектов и серверов, скрипт может работать 1–3 минуты.

- API Selectel использует IAM-токены. Токен действует ограниченное время
  (обычно 1 час). Если токен истёк — создайте новый в панели управления.

---

## Troubleshooting

| Ошибка | Решение |
|--------|---------|
| `401 Unauthorized` | Токен истёк или неверный — создайте новый |
| `403 Forbidden` | Токен не имеет нужных прав — проверьте роль пользователя |
| `Облачные серверы: 0` | Не удалось получить Keystone-токен — используйте `--keystone-token` |
| `Connection error` | Проверьте интернет-соединение |
