"""
MRP-CRP планировщик производства для металлургического предприятия.

Скрипт решает учебную задачу планирования производства:

1. Создаёт и наполняет базу данных PostgreSQL.
2. Формирует аналитические отчёты по сбытовым заказам.
3. Выполняет MRP-CRP расчёт:
   - MRP / EXPLODING: рассчитывает потребность по операциям маршрута;
   - CRP / SHIFTING: распределяет операции по календарю мощностей.
4. Формирует итоговые отчёты:
   - MRP: потребность в материалах;
   - CRP: загрузка мощностей;
   - MRP-CRP: детальный операционный отчёт;
   - Смешанный отчёт оставлен отдельно в конце.
5. Выгружает результаты в Excel.
"""

import datetime
import math

import pandas as pd
import psycopg2

# =============================================================================
# КОНФИГУРАЦИЯ БАЗЫ ДАННЫХ
# =============================================================================

DB_CONFIG = {
    "dbname": "scm_laboratory",
    "user": "postgres",
    "password": "123",
    "port": "5432",
    "host": "localhost",
}

# =============================================================================
# СХЕМА БАЗЫ ДАННЫХ
# =============================================================================

DDL = """
DROP TABLE IF EXISTS MRP_Plan CASCADE;
DROP TABLE IF EXISTS calendar CASCADE;
DROP TABLE IF EXISTS standard_operations CASCADE;
DROP TABLE IF EXISTS sales_orders CASCADE;
DROP TABLE IF EXISTS resources CASCADE;
DROP TABLE IF EXISTS products CASCADE;
DROP TABLE IF EXISTS customers CASCADE;

CREATE TABLE customers (
    customer_id SERIAL PRIMARY KEY,
    customer_group_id INTEGER,
    customer_name VARCHAR(80)
);

CREATE TABLE products (
    product_id SERIAL PRIMARY KEY,
    product_desc VARCHAR(150) NOT NULL,
    product_min_weight FLOAT,
    product_max_weight FLOAT,
    product_group VARCHAR(50),
    product_type VARCHAR(50)
);

CREATE TABLE resources (
    resource_id SERIAL PRIMARY KEY,
    resource_desc VARCHAR(100) NOT NULL,
    wearout INTEGER DEFAULT 0
);

CREATE TABLE sales_orders (
    sales_order_id INTEGER PRIMARY KEY,
    customer_id INTEGER REFERENCES customers(customer_id),
    product_id INTEGER REFERENCES products(product_id),
    target_weight FLOAT NOT NULL,
    tolerance FLOAT NOT NULL,
    unit_weight FLOAT NOT NULL,
    due_date DATE NOT NULL,
    priority INTEGER NOT NULL,
    status INTEGER DEFAULT 1
);

CREATE TABLE standard_operations (
    operation_ref SERIAL PRIMARY KEY,
    product_id INTEGER REFERENCES products(product_id),
    resource_id INTEGER REFERENCES resources(resource_id),
    alternate_pref INTEGER DEFAULT 1,
    operation_id INTEGER DEFAULT 1,
    performance FLOAT,
    yield FLOAT DEFAULT 1.0
);

CREATE TABLE calendar (
    production_date DATE NOT NULL,
    resource_id INTEGER REFERENCES resources(resource_id),
    available_hours FLOAT DEFAULT 23.0,
    PRIMARY KEY (production_date, resource_id)
);

CREATE TABLE MRP_Plan (
    mrp_id SERIAL PRIMARY KEY,
    sales_order_id INTEGER,
    customer_name VARCHAR(150),
    product_desc VARCHAR(150),
    resource_id INTEGER,
    resource_desc VARCHAR(100),
    production_date DATE,
    quantity INTEGER,
    weight FLOAT,
    capacity_hours_required FLOAT,
    status VARCHAR(20) DEFAULT 'planned'
);
"""

# =============================================================================
# ИСХОДНЫЕ ДАННЫЕ
# =============================================================================

PRODUCTS_DATA = [
    {"product_desc": "Лист г/к", "product_group": "лист г/к", "product_type": "hot"},
    {"product_desc": "Рулон г/к", "product_group": "рулон г/к", "product_type": "hot"},

    {"product_desc": "Лист г/к травленный", "product_group": "лист г/к травл.", "product_type": "hot_pickled"},
    {"product_desc": "Лист г/к травл.", "product_group": "лист г/к травл.", "product_type": "hot_pickled"},
    {"product_desc": "Рулон г/к травл.", "product_group": "рулон г/к травл.", "product_type": "hot_pickled"},
    {"product_desc": "Рулон г/к травленный", "product_group": "рулон г/к травл.", "product_type": "hot_pickled"},

    {"product_desc": "Рулон х/к", "product_group": "рулон х/к", "product_type": "cold"},
    {"product_desc": "Лист х/к", "product_group": "лист х/к", "product_type": "cold"},
]

CUSTOMERS_DATA = [
    {"customer_name": "Росатом", "customer_group_id": 1},
    {"customer_name": "ТМК", "customer_group_id": 1},
    {"customer_name": "ОМК", "customer_group_id": 2},
    {"customer_name": "УВЗ", "customer_group_id": 3},
    {"customer_name": "ММЗ", "customer_group_id": 3},
    {"customer_name": "ГАЗ", "customer_group_id": 4},
    {"customer_name": "КрАЗ", "customer_group_id": 4},
]

ORDERS_DATA = [
    {
        "so_item_id": "BS1001",
        "customer": "Росатом",
        "material": "Лист г/к",
        "sku_weight_tons": 12,
        "volume_tons": 2700,
        "tolerance": 15,
        "due_date": "2025-04-15",
        "priority": 1,
    },
    {
        "so_item_id": "BS1002",
        "customer": "Росатом",
        "material": "Лист г/к травленный",
        "sku_weight_tons": 8,
        "volume_tons": 1030,
        "tolerance": 10,
        "due_date": "2025-04-17",
        "priority": 1,
    },
    {
        "so_item_id": "BS1003",
        "customer": "ОМК",
        "material": "Рулон х/к",
        "sku_weight_tons": 15,
        "volume_tons": 1500,
        "tolerance": 30,
        "due_date": "2025-04-28",
        "priority": 2,
    },
    {
        "so_item_id": "BS1004",
        "customer": "ТМК",
        "material": "Лист г/к травл.",
        "sku_weight_tons": 5,
        "volume_tons": 2500,
        "tolerance": 10,
        "due_date": "2025-04-30",
        "priority": 1,
    },
    {
        "so_item_id": "BS1005",
        "customer": "УВЗ",
        "material": "Рулон г/к травл.",
        "sku_weight_tons": 8,
        "volume_tons": 2000,
        "tolerance": 10,
        "due_date": "2025-04-21",
        "priority": 3,
    },
    {
        "so_item_id": "BS1006",
        "customer": "ГАЗ",
        "material": "Рулон х/к",
        "sku_weight_tons": 10,
        "volume_tons": 5800,
        "tolerance": 20,
        "due_date": "2025-04-10",
        "priority": 4,
    },
    {
        "so_item_id": "BS1007",
        "customer": "ММЗ",
        "material": "Рулон г/к травленный",
        "sku_weight_tons": 15,
        "volume_tons": 1800,
        "tolerance": 10,
        "due_date": "2025-04-03",
        "priority": 3,
    },
    {
        "so_item_id": "BS1008",
        "customer": "КрАЗ",
        "material": "Лист х/к",
        "sku_weight_tons": 12,
        "volume_tons": 4000,
        "tolerance": 20,
        "due_date": "2025-04-20",
        "priority": 4,
    },
]

RESOURCES_DATA = [
    {"resource_desc": "Прокатный стан"},
    {"resource_desc": "Агрегат резки1"},
    {"resource_desc": "Линия упаковки1"},
    {"resource_desc": "Агрегат травления"},
    {"resource_desc": "Линия дрессировки"},
    {"resource_desc": "Агрегат резки2"},
    {"resource_desc": "Линия упаковки2"},
]

# Производительность агрегатов, тонн/час
PERFORMANCE = {
    "Прокатный стан": 180.0,
    "Агрегат резки1": 200.0,
    "Линия упаковки1": 150.0,
    "Агрегат травления": 130.0,
    "Линия дрессировки": 150.0,
    "Агрегат резки2": 150.0,
    "Линия упаковки2": 110.0,
}

# Коэффициенты выхода годного.
YIELD_RATES = {
    "Прокатный стан": 0.984,
    "Агрегат резки1": 0.984,
    "Линия упаковки1": 1.000,
    "Агрегат травления": 1.000,
    "Линия дрессировки": 0.992,
    "Агрегат резки2": 0.984,
    "Линия упаковки2": 1.000,
}

# Маршруты по типам продукции
# Резка1 и Упаковка1 используются только для горячекатаного продукта
ROUTES = {
    "hot": [
        "Прокатный стан",
        "Агрегат резки1",
        "Линия упаковки1",
    ],
    "hot_pickled": [
        "Прокатный стан",
        "Агрегат травления",
        "Агрегат резки2",
        "Линия упаковки2",
    ],
    "cold": [
        "Прокатный стан",
        "Агрегат травления",
        "Линия дрессировки",
        "Агрегат резки2",
        "Линия упаковки2",
    ],
}

# Календарные исключения
CALENDAR_EXCEPTIONS = {
    "Прокатный стан": {
        12: 15.0,
        20: 0.0,
        21: 0.0,
    },
    "Агрегат травления": {
        10: 15.0,
        11: 10.0,
        25: 10.0,
    },
}

# Даты начала и конца планирования
PLAN_START = datetime.date(2025, 4, 1)
PLAN_END = datetime.date(2025, 4, 30)

# Дефолтные настройки для планирования по датам
DEFAULT_DAILY_HOURS = 23.0
MAX_ORDER_HOURS_PER_DAY = 24.0
EPSILON = 1e-9

def get_available_hours(resource_name, day):
    """
    Возвращает доступное время агрегата для заданного дня месяца.
    Если для агрегата и дня есть исключение в CALENDAR_EXCEPTIONS,
    возвращается значение из исключения, иначе возвращается 23 часа.
    """
    exceptions = CALENDAR_EXCEPTIONS.get(resource_name, {})
    return exceptions.get(day, DEFAULT_DAILY_HOURS)

def round_to_nearest_units(target_weight, unit_weight):
    """
    Округляет количество единиц продукции до ближайшего целого.
    """
    return math.floor(target_weight / unit_weight + 0.5)

# =============================================================================
# ИНИЦИАЛИЗАЦИЯ И НАПОЛНЕНИЕ БАЗЫ ДАННЫХ
# =============================================================================

def init_and_populate_db():
    """
    Создаёт таблицы и заполняет их исходными данными.
    """

    # Подключение к БД
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
    except Exception as error:
        print(f"[ERROR] Не удалось подключиться к базе данных: {error}")
        return

    # Исполнение DDL-скрипта
    try:
        cursor.execute(DDL)
        conn.commit()

        # Словари для преобразований
        customer_name_to_id = {}
        product_desc_to_id = {}
        resource_desc_to_id = {}

        # Вставка данных заказчиков (с возвращением для дальнейшей работы)
        for customer in CUSTOMERS_DATA:
            cursor.execute(
                """
                INSERT INTO customers (customer_group_id, customer_name)
                VALUES (%s, %s)
                RETURNING customer_id;
                """,
                (
                    customer["customer_group_id"],
                    customer["customer_name"],
                ),
            )

            # Запоминание id клиента
            customer_id = cursor.fetchone()[0]
            customer_name_to_id[customer["customer_name"]] = customer_id

        # Заполнение данных продукции
        for product in PRODUCTS_DATA:
            cursor.execute(
                """
                INSERT INTO products
                    (product_desc, product_group, product_type, product_min_weight, product_max_weight)
                VALUES
                    (%s, %s, %s, 5.0, 25.0)
                RETURNING product_id;
                """,
                (
                    product["product_desc"],
                    product["product_group"],
                    product["product_type"],
                ),
            )

            # Запоминание id продукта
            product_id = cursor.fetchone()[0]
            product_desc_to_id[product["product_desc"]] = product_id

        # Заполнение данных агрегатов
        for resource in RESOURCES_DATA:
            cursor.execute(
                """
                INSERT INTO resources (resource_desc)
                VALUES (%s)
                RETURNING resource_id;
                """,
                (resource["resource_desc"],),
            )

            # Запоминание id агрегата
            resource_id = cursor.fetchone()[0]
            resource_desc_to_id[resource["resource_desc"]] = resource_id

        # Заполнение standard_operations
        for product in PRODUCTS_DATA:
            # Получаем числовой ID продукта из словаря-кэша
            product_id = product_desc_to_id[product["product_desc"]]

            # Извлекаем тип продукта (например, "hot", "cold", "hot_pickled") для определения маршрута
            product_type = product["product_type"]

            route = ROUTES.get(product_type, [])

            for operation_number, machine_name in enumerate(route, start=1):
                # Получаем числовой ID ресурса (станка) из словаря-кэша для преобразования в id
                resource_id = resource_desc_to_id[machine_name]

                # Заполнение мощности и выхода
                performance = PERFORMANCE[machine_name]
                yield_rate = YIELD_RATES[machine_name]

                cursor.execute(
                    """
                    INSERT INTO standard_operations
                        (product_id, resource_id, operation_id, performance, yield)
                    VALUES
                        (%s, %s, %s, %s, %s);
                    """,
                    (
                        product_id,
                        resource_id,
                        operation_number,
                        performance,
                        yield_rate,
                    ),
                )

        # Заполнение календаря доступных мощностей
        # На каждый день горизонта и на каждый агрегат создаётся отдельная строка
        horizon_days = (PLAN_END - PLAN_START).days + 1

        for day_offset in range(horizon_days):

            # Извлекаем номер дня в месяце (1-30 для апреля, 1-31 для других месяцев)
            # Нужен для проверки исключений в календаре
            current_date = PLAN_START + datetime.timedelta(days=day_offset)
            day_number = current_date.day

            for resource_name, resource_id in resource_desc_to_id.items():
                # Получаем количество доступных часов для данного ресурса в конкретный день.
                available_hours = get_available_hours(resource_name, day_number)

                cursor.execute(
                    """
                    INSERT INTO calendar
                        (production_date, resource_id, available_hours)
                    VALUES
                        (%s, %s, %s);
                    """,
                    (
                        current_date,
                        resource_id,
                        available_hours,
                    ),
                )

        # Заполнение сбытовых заказов
        for order in ORDERS_DATA:
            # Преобразование из имени клиента в id
            customer_id = customer_name_to_id[order["customer"]]

            # Преобразование из названия продукта в id
            product_id = product_desc_to_id[order["material"]]

            # Преобразование id в числовое значение
            sales_order_id = int(order["so_item_id"].replace("BS", ""))

            cursor.execute(
                """
                INSERT INTO sales_orders
                    (
                        sales_order_id,
                        customer_id,
                        product_id,
                        target_weight,
                        tolerance,
                        unit_weight,
                        due_date,
                        priority,
                        status
                    )
                VALUES
                    (%s, %s, %s, %s, %s, %s, %s, %s, 1);
                """,
                (
                    sales_order_id,
                    customer_id,
                    product_id,
                    float(order["volume_tons"]),
                    float(order["tolerance"]),
                    float(order["sku_weight_tons"]),
                    order["due_date"],
                    order["priority"],
                ),
            )

        conn.commit()
        print("[OK] База данных инициализирована и наполнена.")

    except Exception as error:
        conn.rollback()
        print(f"[ERROR] Сбой при заполнении таблиц: {error}")
        raise

    finally:
        cursor.close()
        conn.close()

# =============================================================================
# АНАЛИЗ СБЫТОВЫХ ЗАКАЗОВ
# =============================================================================

def show_part1_analytical_reports():
    """
    Формирует аналитические отчёты по сбытовым заказам:
    1. по датам сдачи;
    2. по видам продукции;
    3. по заказчикам.
    """

    conn = psycopg2.connect(**DB_CONFIG)

    # Настройка для вывода таблиц в pandas
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 1200)

    print("\n" + "=" * 80)
    print("АНАЛИЗ СБЫТОВЫХ ЗАКАЗОВ: ОТЧЕТЫ К ЧАСТИ №1")
    print("=" * 80)

    queries = {
        "1. Распределение заказов по датам сдачи": """
            SELECT
                due_date AS "Дата сдачи",
                COUNT(*) AS "Кол-во заказов",
                SUM(target_weight) AS "Общий объем (тонн)"
            FROM sales_orders
            GROUP BY due_date
            ORDER BY due_date;
        """,

        "2. Распределение по видам продукции и объемам": """
            SELECT
                p.product_desc AS "Вид продукции",
                COUNT(so.sales_order_id) AS "Кол-во заказов",
                SUM(so.target_weight) AS "Общий объем (тонн)"
            FROM sales_orders so
            JOIN products p ON so.product_id = p.product_id
            GROUP BY p.product_desc
            ORDER BY SUM(so.target_weight) DESC;
        """,

        "3. Распределение по заказчикам и объемам": """
            SELECT
                c.customer_name AS "Заказчик",
                COUNT(so.sales_order_id) AS "Кол-во заказов",
                SUM(so.target_weight) AS "Общий объем (тонн)"
            FROM sales_orders so
            JOIN customers c ON so.customer_id = c.customer_id
            GROUP BY c.customer_name
            ORDER BY SUM(so.target_weight) DESC;
        """,
    }

    for title, query in queries.items():
        print(f"\n{title}:")
        df = pd.read_sql_query(query, conn)
        print(df.to_string(index=False))

    conn.close()

# =============================================================================
# ОСНОВНОЙ АЛГОРИТМ MRP-CRP
# =============================================================================

def generate_mrp_plan():
    """
    Выполняет основной алгоритм MRP-CRP.
    Алгоритм состоит из двух частей:

    1. MRP / EXPLODING.
       Для каждого заказа рассчитывается потребность по операциям маршрута.
       Расчёт идёт обратным ходом: от упаковки к первой операции.

    2. CRP / SHIFTING.
       Рассчитанный фонд времени распределяется по календарю мощностей.
       Сначала работа размещается назад от даты сдачи, а если мощности
       не хватает — вперёд со статусом "опоздание".

    Ограничения:
    - агрегат не может быть загружен больше, чем доступно в calendar;
    - один заказ не может занимать больше 24 часов в сутки суммарно
      по всем агрегатам.
    """

    conn = psycopg2.connect(**DB_CONFIG)
    cursor = conn.cursor()

    try:
        print("\nЗапуск алгоритма MRP-CRP (Exploding -> Shifting)...")

        # Заказы сортируются по приоритету и дате сдачи
        cursor.execute(
            """
            SELECT
                so.sales_order_id,
                c.customer_name,
                p.product_id,
                p.product_desc,
                so.target_weight,
                so.tolerance,
                so.unit_weight,
                so.due_date
            FROM sales_orders so
            JOIN customers c ON so.customer_id = c.customer_id
            JOIN products p ON so.product_id = p.product_id
            ORDER BY so.priority ASC, so.due_date ASC;
            """
        )

        # Забираем результат запроса выше
        orders = cursor.fetchall()

        for order_row in orders:
            (
                sales_order_id,
                customer_name,
                product_id,
                product_desc,
                target_weight,
                tolerance,
                unit_weight,
                due_date,
            ) = order_row

            # -----------------------------------------------------------------
            # MRP. Шаг 1. Плановый вес сдачи
            # -----------------------------------------------------------------
            # Заказ должен быть выполнен целым числом рулонов/пачек
            # Поэтому целевой вес округляется к ближайшему кратному unit_weight
            planned_qty = round_to_nearest_units(target_weight, unit_weight)
            planned_weight = planned_qty * unit_weight

            deviation = abs(target_weight - planned_weight)

            if deviation > tolerance:
                print(
                    f"  [WARN] Заказ BS{sales_order_id}: "
                    f"отклонение {deviation:.2f} т превышает допуск +/-{tolerance} т. "
                    f"Принят ближайший кратный вес: {planned_weight:.2f} т."
                )

            # -----------------------------------------------------------------
            # MRP. Шаг 2. Получение технологического маршрута
            # -----------------------------------------------------------------
            cursor.execute(
                """
                SELECT
                    op.resource_id,
                    r.resource_desc,
                    op.performance,
                    op.yield,
                    op.operation_id
                FROM standard_operations op
                JOIN resources r ON op.resource_id = r.resource_id
                WHERE op.product_id = %s
                ORDER BY op.operation_id ASC;
                """,
                (product_id,),
            )

            operations = cursor.fetchall()

            if not operations:
                print(
                    f"  [WARN] Заказ BS{sales_order_id}: маршрут не найден, заказ пропущен."
                )
                continue

            # -----------------------------------------------------------------
            # MRP. Шаг 3. EXPLODING
            # -----------------------------------------------------------------
            # operations хранит маршрут в прямом порядке:
            #   Прокат -> ... -> Упаковка
            #
            # Для разузлования нужен обратный ход:
            #   Упаковка -> ... -> Прокат
            #
            # На последней операции вес равен плановому весу сдачи.
            # На предыдущих операциях вес увеличивается с учётом выхода годного.
            exploded_tasks = []

            # Текущий вес и количество на выходе из последней операции
            current_weight = planned_weight
            current_qty = planned_qty

            for index, operation_row in enumerate(reversed(operations)):
                (
                    resource_id,
                    resource_desc,
                    performance,
                    yield_rate,
                    operation_id,
                ) = operation_row

                if index > 0:
                    # Чтобы после потерь получить нужный выход на вход операции нужно подать больший вес
                    current_weight = current_weight / yield_rate

                    # Полуфабрикат тоже должен быть кратен весу единицы
                    # Здесь округляем вверх, чтобы материала точно хватило
                    current_qty = math.ceil(current_weight / unit_weight)
                    current_weight = current_qty * unit_weight

                # Требуемый фонд времени операции.
                hours_required = current_weight / performance

                # Добавляем в список задач для CRP-планирования словарь с параметрами операции
                # Эти данные позже будут использованы для распределения загрузки по календарю
                exploded_tasks.append(
                    {
                        "resource_id": resource_id,
                        "resource_desc": resource_desc,
                        "qty": current_qty,
                        "weight": current_weight,
                        "hours_required": hours_required,
                    }
                )

            print(
                f"\n  Заказ BS{sales_order_id} "
                f"({product_desc}, {planned_weight:.0f} т, сдача {due_date}):"
            )

            for task in exploded_tasks:
                print(
                    f"    {task['resource_desc']:25s} "
                    f"вес={task['weight']:.1f} т "
                    f"фонд={task['hours_required']:.2f} ч"
                )

            # -----------------------------------------------------------------
            # CRP. Шаг 4. SHIFTING
            # -----------------------------------------------------------------
            # Здесь рассчитанные MRP-потребности размещаются по календарю
            # доступных мощностей.
            #
            # order_daily_hours нужен для правила 24:
            # один заказ не может занимать больше 24 часов в сутки суммарно
            # по всем агрегатам.
            order_daily_hours = {}

            # Для упаковки верхняя граница — due_date.
            # Для каждой предыдущей операции граница сдвигается к самой ранней дате, занятой последующей операцией.
            next_op_start_limit = due_date

            # Флаг обрыва цепочки. Если не смогли разместить операцию в горизонте планирования
            chain_broken = False

            # Перебор задач текущего заказа
            for task in exploded_tasks:
                resource_id = task["resource_id"]
                resource_desc = task["resource_desc"]

                # Сколько часов нужно для этой операции (еще не размещенных)
                hours_left = task["hours_required"]

                # Общее количество часов для операции
                total_hours = task["hours_required"]

                # Самая ранняя дата, на которую удалось разместить хоть часть операции
                min_date_used = None

                # -------------------------------------------------------------
                # CRP. Backward shifting: размещение назад от срока
                # -------------------------------------------------------------

                # Начинаем планирование с даты, ограниченной предыдущей операцией
                target_date = next_op_start_limit
                order_status = "Запланирован"

                # Пытаемся разместить все часы (hours_left) в доступные слоты
                while hours_left > EPSILON and target_date >= PLAN_START:
                    # Запрос часов агрегата
                    cursor.execute(
                        """
                        SELECT available_hours
                        FROM calendar
                        WHERE production_date = %s
                          AND resource_id = %s;
                        """,
                        (
                            target_date,
                            resource_id,
                        ),
                    )

                    row = cursor.fetchone()
                    # Получаем результат. Если записи нет (row = None), доступных часов 0
                    available_hours = row[0] if row else 0.0

                    # Сколько часов от этого заказа уже выделено на этот день
                    # order_daily_hours — словарь, отслеживающий загрузку по дням для текущего заказа
                    used_today = order_daily_hours.get(target_date, 0.0)

                    # Сколько еще можно выделить на этот день по правилу
                    allowed_by_rule24 = max(
                        0.0,
                        MAX_ORDER_HOURS_PER_DAY - used_today,
                    )

                    # Основное ограничение CRP:
                    # можно разместить не больше минимума из:
                    # 1) оставшихся часов операции
                    # 2) свободных часов агрегата
                    # 3) остатка лимита 24 часа по заказу на этот день
                    allocated_hours = min(
                        hours_left,
                        available_hours,
                        allowed_by_rule24,
                    )

                    # Если удалось выделить хоть сколько-то часов
                    if allocated_hours > EPSILON:
                        # Вычисляем, какую долю от всей операции мы сейчас разместили
                        ratio = allocated_hours / total_hours

                        # Вес, который пройдет через операцию в этот день, пропорционален часам
                        daily_weight = task["weight"] * ratio

                        # Количество единиц продукции, округленное пропорционально
                        daily_qty = round(task["qty"] * ratio)

                        # Защита от нуля: если должны сделать хоть что-то, но получилось 0,
                        # устанавливаем минимум 1 единицу
                        if daily_qty == 0 and task["qty"] > 0:
                            daily_qty = 1

                        # Вставляем запись в MRP_Plan — детальный операционный план
                        cursor.execute(
                            """
                            INSERT INTO MRP_Plan
                                (
                                    sales_order_id,
                                    customer_name,
                                    product_desc,
                                    resource_id,
                                    resource_desc,
                                    production_date,
                                    quantity,
                                    weight,
                                    capacity_hours_required,
                                    status
                                )
                            VALUES
                                (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
                            """,
                            (
                                sales_order_id,
                                customer_name,
                                product_desc,
                                resource_id,
                                resource_desc,
                                target_date,
                                daily_qty,
                                round(daily_weight, 2),
                                round(allocated_hours, 4),
                                order_status,
                            ),
                        )

                        # Уменьшаем доступную мощность агрегата
                        # Cледующий заказ увидит уже занятые часы
                        cursor.execute(
                            """
                            UPDATE calendar
                            SET available_hours = available_hours - %s
                            WHERE production_date = %s
                              AND resource_id = %s;
                            """,
                            (
                                allocated_hours,
                                target_date,
                                resource_id,
                            ),
                        )

                        # Уменьшаем оставшиеся часы операции на то, что разместили
                        hours_left -= allocated_hours

                        # Обновляем словарь использованных часов для этого заказа
                        order_daily_hours[target_date] = used_today + allocated_hours

                        # Запоминаем самую раннюю дату, где разместили часть операции
                        # Она станет ограничителем для следующей (предыдущей в маршруте) операции
                        if min_date_used is None or target_date < min_date_used:
                            min_date_used = target_date

                    # Переходим к предыдущиму дню
                    target_date -= datetime.timedelta(days=1)

                # -------------------------------------------------------------
                # CRP. Forward shifting: если назад разместить не получилось
                # Это означает что между due_date и PLAN_START не хватило мощности станка
                # -------------------------------------------------------------

                # Остались ли часы, которые не удалось разместить при движении назад
                if hours_left > EPSILON:
                    target_date = next_op_start_limit + datetime.timedelta(days=1)
                    order_status = "опоздание"

                    # Пытаемся разместить оставшиеся часы, двигаясь вперед по календарю
                    while hours_left > EPSILON and target_date <= PLAN_END:

                        #  Запрашиваем доступные часы станка в текущую дату
                        cursor.execute(
                            """
                            SELECT available_hours
                            FROM calendar
                            WHERE production_date = %s
                              AND resource_id = %s;
                            """,
                            (
                                target_date,
                                resource_id,
                            ),
                        )

                        row = cursor.fetchone()
                        available_hours = row[0] if row else 0.0

                        # Сколько часов от этого заказа уже выделено на этот день
                        used_today = order_daily_hours.get(target_date, 0.0)

                        # Ограничение: не более 24 часов на заказ в день
                        allowed_by_rule24 = max(
                            0.0,
                            MAX_ORDER_HOURS_PER_DAY - used_today,
                        )

                        # Сколько можем выделить сегодня? Минимум из:
                        # - оставшихся часов операции
                        # - свободных часов станка
                        # - остатка лимита 24 часа
                        allocated_hours = min(
                            hours_left,
                            available_hours,
                            allowed_by_rule24,
                        )

                        # Если удалось выделить хоть сколько-то часов
                        if allocated_hours > EPSILON:
                            # Доля от всей операции, которую размещаем сегодня
                            ratio = allocated_hours / total_hours

                            # Пропорционально распределяем вес и количество
                            daily_weight = task["weight"] * ratio
                            daily_qty = round(task["qty"] * ratio)

                            # Защита от нуля: если должны сделать хоть что-то, ставим 1 единицу
                            if daily_qty == 0 and task["qty"] > 0:
                                daily_qty = 1

                            # Сохраняем детальную запись в MRP_Plan
                            cursor.execute(
                                """
                                INSERT INTO MRP_Plan
                                    (
                                        sales_order_id,
                                        customer_name,
                                        product_desc,
                                        resource_id,
                                        resource_desc,
                                        production_date,
                                        quantity,
                                        weight,
                                        capacity_hours_required,
                                        status
                                    )
                                VALUES
                                    (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
                                """,
                                (
                                    sales_order_id,
                                    customer_name,
                                    product_desc,
                                    resource_id,
                                    resource_desc,
                                    target_date,
                                    daily_qty,
                                    round(daily_weight, 2),
                                    round(allocated_hours, 4),
                                    order_status,
                                ),
                            )

                            # Уменьшаем мощность агрегата
                            cursor.execute(
                                """
                                UPDATE calendar
                                SET available_hours = available_hours - %s
                                WHERE production_date = %s
                                  AND resource_id = %s;
                                """,
                                (
                                    allocated_hours,
                                    target_date,
                                    resource_id,
                                ),
                            )

                            # Уменьшаем остаток неразмещенных часов
                            hours_left -= allocated_hours

                            # Обновляем счетчик использованных часов для этого заказа
                            order_daily_hours[target_date] = used_today + allocated_hours

                            # Запоминаем самую раннюю дату, где разместили операцию
                            if min_date_used is None or target_date < min_date_used:
                                min_date_used = target_date

                        # Переходим к следующему дню
                        target_date += datetime.timedelta(days=1)

                # -------------------------------------------------------------
                # Передача даты следующей операции
                # -------------------------------------------------------------
                # Если текущая операция была размещена, предыдущая операция
                # должна завершиться не позже самой ранней даты текущей операции.
                if min_date_used is not None:
                    next_op_start_limit = min_date_used
                else:
                    print(
                        f"  [ERROR] Агрегат '{resource_desc}', "
                        f"заказ BS{sales_order_id}: нет доступных часов "
                        f"на всём горизонте планирования."
                    )

                    # Ставим флаг, что цепочка операций для этого заказа прервана
                    chain_broken = True
                    break # Выходим из цикла по операциям (весь заказ невыполним)

            if chain_broken:
                cursor.execute(
                    """
                    UPDATE MRP_Plan
                    SET status = 'ошибка мощности'
                    WHERE sales_order_id = %s;
                    """,
                    (sales_order_id,),
                )

        conn.commit()
        print("\n[OK] Алгоритм MRP-CRP завершён.")

    except Exception as error:
        # Откат изменений в случае ошибки
        conn.rollback()
        print(f"[ERROR] Критический сбой алгоритма: {error}")
        raise

    finally:
        cursor.close()
        conn.close()

# =============================================================================
# ОТЧЁТЫ: CRP, MRP И ДЕТАЛЬНЫЙ MRP-CRP
# =============================================================================

def show_required_capacity_report():
    """
    CRP-отчёт по требуемой мощности.
    Показывает, сколько часов требуется каждому заказу
    на каждом агрегате по датам
    """

    conn = psycopg2.connect(**DB_CONFIG)

    query = """
        SELECT
            sales_order_id AS "ID Заказа",
            customer_name AS "Заказчик",
            production_date AS "Дата планирования",
            resource_desc AS "Агрегат",
            SUM(capacity_hours_required) AS "CRP: Мощность (ч)",
            status AS "Статус"
        FROM MRP_Plan
        GROUP BY
            sales_order_id,
            customer_name,
            production_date,
            resource_desc,
            status
        ORDER BY
            sales_order_id,
            production_date,
            resource_desc;
    """

    print("\n" + "=" * 80)
    print("CRP-ОТЧЁТ: ТРЕБУЕМАЯ МОЩНОСТЬ ПО ЗАКАЗАМ / АГРЕГАТАМ / ДАТАМ")
    print("=" * 80)

    df = pd.read_sql_query(query, conn)
    print(df.to_string(index=False))

    conn.close()

def show_final_mrp_report():
    """
    Формирует итоговые отчёты после расчёта MRP-CRP
    - MRP-отчёт показывает потребность в материалах без агрегатов
    - CRP-отчёт показывает загрузку мощностей
    - MRP-CRP-отчёт показывает детальную связь материала, операции и агрегата
    - Смешанный отчёт оставлен в конце для совместимости с предыдущей версией
    """

    conn = psycopg2.connect(**DB_CONFIG)

    # -------------------------------------------------------------------------
    # MRP: чистая потребность в материалах
    # -------------------------------------------------------------------------
    # Отчёт отвечает на вопрос:
    # "Сколько какого материала требуется по датам производства?"
    mrp_materials_query = """
        SELECT
            production_date AS "MRP: Дата производства",
            product_desc AS "MRP: Материал",
            SUM(weight) AS "MRP: Вес брутто (т)",
            SUM(quantity) AS "MRP: Кол-во (рул/пач)"
        FROM MRP_Plan
        GROUP BY
            production_date,
            product_desc
        ORDER BY
            production_date,
            product_desc;
    """

    # -------------------------------------------------------------------------
    # CRP: сводная загрузка агрегатов
    # -------------------------------------------------------------------------
    # Отчёт отвечает на вопрос:
    # "Сколько часов заняты агрегаты по датам?"
    crp_capacity_summary_query = """
        SELECT
            production_date AS "CRP: Дата",
            resource_desc AS "CRP: Агрегат",
            SUM(capacity_hours_required) AS "CRP: Занято (ч)"
        FROM MRP_Plan
        GROUP BY
            production_date,
            resource_desc
        ORDER BY
            production_date,
            resource_desc;
    """

    # -------------------------------------------------------------------------
    # MRP-CRP: детальный операционный отчёт
    # -------------------------------------------------------------------------
    # Он связывает материал, агрегат, дату операции, вес и часы.
    # Удобен для проверки технологической цепочки.
    operation_detail_query = """
        SELECT
            sales_order_id AS "Заказ",
            production_date AS "Дата операции",
            product_desc AS "Материал",
            resource_desc AS "Агрегат",
            SUM(weight) AS "Вес брутто (т)",
            SUM(quantity) AS "Кол-во (рул/пач)",
            SUM(capacity_hours_required) AS "Мощность (ч)",
            status AS "Статус"
        FROM MRP_Plan
        GROUP BY
            sales_order_id,
            production_date,
            product_desc,
            resource_desc,
            status
        ORDER BY
            sales_order_id,
            production_date,
            resource_desc;
    """

    # -------------------------------------------------------------------------
    # Старый отчёт
    # -------------------------------------------------------------------------
    # Это вид отчёта "Потребность материалов", где материал и агрегат
    # находятся в одной таблице. Формально он не ошибочный, но смешивает MRP и CRP.
    old_mixed_report_query = """
        SELECT
            production_date AS "Дата готовности",
            product_desc AS "Материал",
            resource_desc AS "Агрегат",
            SUM(quantity) AS "Кол-во (рул/пач)",
            SUM(weight) AS "Вес брутто (т)"
        FROM MRP_Plan
        GROUP BY
            production_date,
            product_desc,
            resource_desc
        ORDER BY
            production_date,
            product_desc,
            resource_desc;
    """

    print("\n" + "=" * 80)
    print("MRP-ОТЧЁТ: СВОДНАЯ ПОТРЕБНОСТЬ В МАТЕРИАЛАХ")
    print("=" * 80)
    print(pd.read_sql_query(mrp_materials_query, conn).to_string(index=False))

    print("\n" + "=" * 80)
    print("CRP-ОТЧЁТ: СВОДНАЯ ЗАГРУЗКА АГРЕГАТОВ")
    print("=" * 80)
    print(pd.read_sql_query(crp_capacity_summary_query, conn).to_string(index=False))

    print("\n" + "=" * 80)
    print("MRP-CRP-ОТЧЁТ: ДЕТАЛЬНАЯ ПОТРЕБНОСТЬ ПО ОПЕРАЦИЯМ")
    print("=" * 80)
    print(pd.read_sql_query(operation_detail_query, conn).to_string(index=False))

    print("\n" + "=" * 80)
    print("СТАРЫЙ СМЕШАННЫЙ ОТЧЁТ: МАТЕРИАЛЫ + АГРЕГАТЫ")
    print("=" * 80)
    print(pd.read_sql_query(old_mixed_report_query, conn).to_string(index=False))

    conn.close()

# =============================================================================
# ВЫГРУЗКА В EXCEL
# =============================================================================

def export_all_to_excel(file_name="mrp_crp_results.xlsx"):
    """
    Выгружает основные отчёты в Excel-файл.
    """

    print(f"\nВыгрузка отчётов в файл '{file_name}'...")

    conn = None

    try:
        conn = psycopg2.connect(**DB_CONFIG)

        with pd.ExcelWriter(file_name, engine="openpyxl") as writer:

            # -----------------------------------------------------------------
            # Аналитика по сбытовым заказам
            # -----------------------------------------------------------------

            pd.read_sql_query(
                """
                SELECT
                    due_date AS "Дата сдачи",
                    COUNT(*) AS "Кол-во заказов",
                    SUM(target_weight) AS "Объем (т)"
                FROM sales_orders
                GROUP BY due_date
                ORDER BY due_date;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="Анализ - Даты сдачи",
                index=False,
            )

            pd.read_sql_query(
                """
                SELECT
                    p.product_desc AS "Продукция",
                    COUNT(so.sales_order_id) AS "Кол-во заказов",
                    SUM(so.target_weight) AS "Объем (т)"
                FROM sales_orders so
                JOIN products p ON so.product_id = p.product_id
                GROUP BY p.product_desc
                ORDER BY SUM(so.target_weight) DESC;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="Анализ - Продукция",
                index=False,
            )

            pd.read_sql_query(
                """
                SELECT
                    c.customer_name AS "Заказчик",
                    COUNT(so.sales_order_id) AS "Кол-во заказов",
                    SUM(so.target_weight) AS "Объем (т)"
                FROM sales_orders so
                JOIN customers c ON so.customer_id = c.customer_id
                GROUP BY c.customer_name
                ORDER BY SUM(so.target_weight) DESC;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="Анализ - Заказчики",
                index=False,
            )

            # -----------------------------------------------------------------
            # CRP: мощность по заказам
            # -----------------------------------------------------------------
            # Показывает загрузку агрегатов в разрезе заказов.
            pd.read_sql_query(
                """
                SELECT
                    sales_order_id AS "ID Заказа",
                    customer_name AS "Заказчик",
                    resource_desc AS "CRP: Агрегат",
                    production_date AS "CRP: Дата",
                    SUM(capacity_hours_required) AS "CRP: Мощность (ч)",
                    status AS "Статус"
                FROM MRP_Plan
                GROUP BY
                    sales_order_id,
                    customer_name,
                    resource_desc,
                    production_date,
                    status
                ORDER BY
                    sales_order_id,
                    production_date,
                    resource_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="CRP - Мощности по заказам",
                index=False,
            )

            # -----------------------------------------------------------------
            # CRP: сводная загрузка агрегатов
            # -----------------------------------------------------------------
            # Показывает суммарную занятость каждого агрегата по датам.
            pd.read_sql_query(
                """
                SELECT
                    production_date AS "CRP: Дата",
                    resource_desc AS "CRP: Агрегат",
                    SUM(capacity_hours_required) AS "CRP: Занято (ч)"
                FROM MRP_Plan
                GROUP BY
                    production_date,
                    resource_desc
                ORDER BY
                    production_date,
                    resource_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="CRP - Сводная загрузка",
                index=False,
            )

            # -----------------------------------------------------------------
            # MRP: чистая потребность в материалах
            # -----------------------------------------------------------------
            # Здесь нет агрегата: это именно материал, дата, вес и количество.
            pd.read_sql_query(
                """
                SELECT
                    production_date AS "MRP: Дата производства",
                    product_desc AS "MRP: Материал",
                    SUM(weight) AS "MRP: Вес брутто (т)",
                    SUM(quantity) AS "MRP: Кол-во (рул/пач)"
                FROM MRP_Plan
                GROUP BY
                    production_date,
                    product_desc
                ORDER BY
                    production_date,
                    product_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="MRP - Потребность",
                index=False,
            )

            # -----------------------------------------------------------------
            # MRP-CRP: детальный операционный отчёт
            # -----------------------------------------------------------------
            # Здесь материал связан с агрегатом и мощностью.
            # Это удобно для проверки расчёта, но это уже смешанный отчёт.
            pd.read_sql_query(
                """
                SELECT
                    sales_order_id AS "Заказ",
                    production_date AS "Дата операции",
                    product_desc AS "Материал",
                    resource_desc AS "Агрегат",
                    SUM(weight) AS "Вес брутто (т)",
                    SUM(quantity) AS "Кол-во (рул/пач)",
                    SUM(capacity_hours_required) AS "Мощность (ч)",
                    status AS "Статус"
                FROM MRP_Plan
                GROUP BY
                    sales_order_id,
                    production_date,
                    product_desc,
                    resource_desc,
                    status
                ORDER BY
                    sales_order_id,
                    production_date,
                    resource_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="MRP-CRP - Операции",
                index=False,
            )

            # -----------------------------------------------------------------
            # Календарь: остаток свободных часов
            # -----------------------------------------------------------------
            # Это дополнительная проверка CRP: сколько мощности осталось после
            # размещения всех заказов.
            pd.read_sql_query(
                """
                SELECT
                    c.production_date AS "Дата",
                    r.resource_desc AS "Агрегат",
                    c.available_hours AS "Остаток (ч)"
                FROM calendar c
                JOIN resources r ON c.resource_id = r.resource_id
                ORDER BY
                    c.production_date,
                    r.resource_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="Календарь - Свободные часы",
                index=False,
            )

            # -----------------------------------------------------------------
            # Старый отчёт
            # -----------------------------------------------------------------
            # Оставлен в конце, как отдельный лист.
            # Формально он корректен, но смешивает MRP и CRP:
            # материал + агрегат в одной таблице.
            pd.read_sql_query(
                """
                SELECT
                    production_date AS "Дата готовности",
                    product_desc AS "Материал",
                    resource_desc AS "Агрегат",
                    SUM(quantity) AS "Кол-во (рул/пач)",
                    SUM(weight) AS "Вес брутто (т)"
                FROM MRP_Plan
                GROUP BY
                    production_date,
                    product_desc,
                    resource_desc
                ORDER BY
                    production_date,
                    product_desc,
                    resource_desc;
                """,
                conn,
            ).to_excel(
                writer,
                sheet_name="Материалы с агрегатом",
                index=False,
            )

        print(f"[OK] Файл '{file_name}' создан.")

    except Exception as error:
        print(f"[ERROR] Ошибка выгрузки в Excel: {error}")

    finally:
        if conn is not None:
            conn.close()

# =============================================================================
# ТОЧКА ВХОДА
# =============================================================================

if __name__ == "__main__":
    init_and_populate_db()
    show_part1_analytical_reports()
    generate_mrp_plan()
    show_required_capacity_report()
    show_final_mrp_report()
    export_all_to_excel()
