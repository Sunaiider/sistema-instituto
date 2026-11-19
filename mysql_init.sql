CREATE DATABASE IF NOT EXISTS instituto CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE instituto;

CREATE TABLE IF NOT EXISTS estudiantes (
    id INT NOT NULL AUTO_INCREMENT,
    numero INT NOT NULL,
    modalidad VARCHAR(50) NOT NULL,
    nivel VARCHAR(20) NOT NULL,
    seccion VARCHAR(10) NOT NULL,
    nombres_apellidos VARCHAR(200) NOT NULL,
    sexo ENUM('M', 'F', 'Otro') NOT NULL,
    mat DECIMAL(5,2) DEFAULT NULL,
    lyl DECIMAL(5,2) DEFAULT NULL,
    l_ext DECIMAL(5,2) DEFAULT NULL,
    en_v DECIMAL(5,2) DEFAULT NULL,
    dmm DECIMAL(5,2) DEFAULT NULL,
    eeff DECIMAL(5,2) DEFAULT NULL,
    aep DECIMAL(5,2) DEFAULT NULL,
    fis DECIMAL(5,2) DEFAULT NULL,
    biol DECIMAL(5,2) DEFAULT NULL,
    ccs DECIMAL(5,2) DEFAULT NULL,
    dpc DECIMAL(5,2) DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_numero (numero)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

INSERT INTO estudiantes (
    numero,
    modalidad,
    nivel,
    seccion,
    nombres_apellidos,
    sexo,
    mat,
    lyl,
    l_ext,
    en_v,
    dmm,
    eeff,
    aep,
    fis,
    biol,
    ccs,
    dpc
) VALUES
(1, 'Regular', 'Secundaria', 'A', 'Juan Perez', 'M', 85.00, 90.00, 80.00, 88.00, 75.00, 82.00, 79.00, 84.00, 76.00, 81.00, 87.00),
(2, 'Regular', 'Secundaria', 'A', 'Maria Lopez', 'F', 90.00, 88.00, 86.00, 91.00, 85.00, 80.00, 89.00, 82.00, 87.00, 84.00, 90.00);
