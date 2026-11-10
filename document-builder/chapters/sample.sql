-- Лабораторна 1: створення таблиць
-- Цей коментар не має потрапити у звіт

CREATE TABLE clients (
    id   SERIAL PRIMARY KEY,          -- ідентифікатор
    name VARCHAR(100) NOT NULL,       /* ім'я клієнта */
    note TEXT DEFAULT 'it''s -- not a comment'
);

/* Багаторядковий
   коментар */
CREATE TABLE orders (
    id        SERIAL PRIMARY KEY,
    client_id INT NOT NULL REFERENCES clients(id),
    total     NUMERIC(10, 2) CHECK (total >= 0)
);
-- [END DDL]

INSERT INTO clients (name) VALUES ('Іван');
