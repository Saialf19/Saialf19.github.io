CREATE DATABASE AHPDB;
GO
USE AHPDB;
GO
CREATE TABLE Criterios (
  id INT IDENTITY(1,1) PRIMARY KEY,
  nombre NVARCHAR(100) NOT NULL
);
CREATE TABLE Alternativas (
  id INT IDENTITY(1,1) PRIMARY KEY,
  nombre NVARCHAR(100) NOT NULL
);
