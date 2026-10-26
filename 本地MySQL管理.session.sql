CREATE DATABASE IF NOT EXISTS wenheng
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

CREATE USER IF NOT EXISTS 'wenheng'@'localhost'
  IDENTIFIED BY '123456';

CREATE USER IF NOT EXISTS 'wenheng'@'127.0.0.1'
  IDENTIFIED BY '123456';

GRANT ALL PRIVILEGES ON wenheng.*
  TO 'wenheng'@'localhost';

GRANT ALL PRIVILEGES ON wenheng.*
  TO 'wenheng'@'127.0.0.1';

SHOW DATABASES LIKE 'wenheng';