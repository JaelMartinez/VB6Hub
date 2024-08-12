CREATE TABLE Libros (
    LibroID INT PRIMARY KEY IDENTITY(1,1),
    Titulo NVARCHAR(255) NOT NULL,
    Autor NVARCHAR(255),
    FechaPublicacion DATE,
    Genero NVARCHAR(100),
    Calificacion DECIMAL(3,2),  -- Ejemplo: 4.50
    EstadoLectura NVARCHAR(50)  -- Ejemplo: "Leído", "Por leer", "No gustó"
);
CREATE TABLE Generos (
    GeneroID INT PRIMARY KEY IDENTITY(1,1),
    NombreGenero NVARCHAR(100) NOT NULL
);
CREATE TABLE LibrosRecomendados (
    RecomendacionID INT PRIMARY KEY IDENTITY(1,1),
    LibroID INT FOREIGN KEY REFERENCES Libros(LibroID),
    FechaRecomendacion DATE
);
CREATE TABLE LibrosLeidos (
    LibroLeidoID INT PRIMARY KEY IDENTITY(1,1),
    LibroID INT FOREIGN KEY REFERENCES Libros(LibroID),
    FechaLectura DATE
);
CREATE TABLE LibrosQueQuieresLeer (
    QuieresLeerID INT PRIMARY KEY IDENTITY(1,1),
    LibroID INT FOREIGN KEY REFERENCES Libros(LibroID),
    Prioridad INT  -- Puede ser 1 para alta, 2 para media, etc.
);
CREATE TABLE LibrosQueNoTeGustan (
    NoTeGustaID INT PRIMARY KEY IDENTITY(1,1),
    LibroID INT FOREIGN KEY REFERENCES Libros(LibroID),
    Motivo NVARCHAR(255)
);

select * from Libros

INSERT INTO Libros (Titulo, Autor, FechaPublicacion, Genero, Calificacion, EstadoLectura) 
VALUES 
('El Hobbit', 'J.R.R. Tolkien', '1937-09-21', 'Fantasía', 4.8, 'Leído'),
('Harry Potter y la piedra filosofal', 'J.K. Rowling', '1997-06-26', 'Fantasía', 4.7, 'Por leer'),
('El nombre del viento', 'Patrick Rothfuss', '2007-03-27', 'Fantasía', 4.5, 'Leído'),
('Juego de tronos', 'George R.R. Martin', '1996-08-06', 'Fantasía', 4.6, 'No gustó'),
('La espada de Shannara', 'Terry Brooks', '1977-02-12', 'Fantasía', 4.0, 'Leído'),
('El último deseo', 'Andrzej Sapkowski', '1993-05-15', 'Fantasía', 4.5, 'Por leer');

INSERT INTO Libros (Titulo, Autor, FechaPublicacion, Genero, Calificacion, EstadoLectura) 
VALUES 
('Dune', 'Frank Herbert', '1965-08-01', 'Ciencia Ficción', 4.7, 'Leído'),
('Neuromante', 'William Gibson', '1984-07-01', 'Ciencia Ficción', 4.3, 'Por leer'),
('1984', 'George Orwell', '1949-06-08', 'Ciencia Ficción', 4.8, 'Leído'),
('Fundación', 'Isaac Asimov', '1951-05-02', 'Ciencia Ficción', 4.6, 'Leído'),
('Snow Crash', 'Neal Stephenson', '1992-06-01', 'Ciencia Ficción', 4.2, 'No gustó'),
('La máquina del tiempo', 'H.G. Wells', '1895-05-07', 'Ciencia Ficción', 4.3, 'Por leer');
 
 INSERT INTO Libros (Titulo, Autor, FechaPublicacion, Genero, Calificacion, EstadoLectura) 
VALUES 
('El código Da Vinci', 'Dan Brown', '2003-03-18', 'Misterio', 4.1, 'No gustó'),
('Sherlock Holmes: El sabueso de los Baskerville', 'Arthur Conan Doyle', '1902-04-01', 'Misterio', 4.7, 'Leído'),
('Los hombres que no amaban a las mujeres', 'Stieg Larsson', '2005-08-18', 'Misterio', 4.5, 'Por leer'),
('La chica del tren', 'Paula Hawkins', '2015-01-13', 'Misterio', 4.0, 'Leído'),
('El silencio de los corderos', 'Thomas Harris', '1988-05-23', 'Misterio', 4.6, 'Leído'),
('Cementerio de animales', 'Stephen King', '1983-11-14', 'Misterio', 4.4, 'Por leer');

INSERT INTO Libros (Titulo, Autor, FechaPublicacion, Genero, Calificacion, EstadoLectura) 
VALUES 
('Orgullo y prejuicio', 'Jane Austen', '1813-01-28', 'Romance', 4.8, 'Leído'),
('Bajo la misma estrella', 'John Green', '2012-01-10', 'Romance', 4.5, 'Por leer'),
('Jane Eyre', 'Charlotte Brontë', '1847-10-16', 'Romance', 4.7, 'Leído'),
('Cumbres borrascosas', 'Emily Brontë', '1847-12-01', 'Romance', 4.4, 'No gustó'),
('Un paseo para recordar', 'Nicholas Sparks', '1999-10-01', 'Romance', 4.2, 'Por leer'),
('Lo que el viento se llevó', 'Margaret Mitchell', '1936-06-30', 'Romance', 4.5, 'Leído');

INSERT INTO Libros (Titulo, Autor, FechaPublicacion, Genero, Calificacion, EstadoLectura) 
VALUES 
('It', 'Stephen King', '1986-09-15', 'Terror', 4.6, 'Leído'),
('Drácula', 'Bram Stoker', '1897-05-26', 'Terror', 4.5, 'Leído'),
('El exorcista', 'William Peter Blatty', '1971-08-01', 'Terror', 4.4, 'Por leer'),
('Frankenstein', 'Mary Shelley', '1818-01-01', 'Terror', 4.7, 'Leído'),
('El resplandor', 'Stephen King', '1977-01-28', 'Terror', 4.8, 'Leído'),
('La llamada de Cthulhu', 'H.P. Lovecraft', '1928-08-01', 'Terror', 4.6, 'Por leer');

select * from LibrosQueNoTeGustan
ALTER TABLE Libros
ADD Sinopsis NVARCHAR(MAX);

SELECT TOP 2 L.Genero, COUNT(L.LibroID) AS TotalLeidos
FROM LibrosLeidos LL
INNER JOIN Libros L ON LL.LibroID = L.LibroID
GROUP BY L.Genero
ORDER BY TotalLeidos DESC;
