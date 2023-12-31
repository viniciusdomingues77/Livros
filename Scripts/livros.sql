USE [Livros]
GO
/****** Object:  User [user_livro]    Script Date: 29/11/2023 21:03:54 ******/
CREATE USER [user_livro] FOR LOGIN [user_livro] WITH DEFAULT_SCHEMA=[dbo]
GO
/****** Object:  User [vini]    Script Date: 29/11/2023 21:03:54 ******/
CREATE USER [vini] FOR LOGIN [vini] WITH DEFAULT_SCHEMA=[dbo]
GO
ALTER ROLE [db_owner] ADD MEMBER [user_livro]
GO
ALTER ROLE [db_owner] ADD MEMBER [vini]
GO
/****** Object:  Table [dbo].[LivroAssunto]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LivroAssunto](
	[Livro_CodL] [int] NOT NULL,
	[Assunto_CodAs] [int] NOT NULL,
 CONSTRAINT [PK_LivroAssunto] PRIMARY KEY CLUSTERED 
(
	[Livro_CodL] ASC,
	[Assunto_CodAs] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Assunto]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Assunto](
	[CodAs] [int] IDENTITY(1,1) NOT NULL,
	[Descricao] [varchar](40) NOT NULL,
 CONSTRAINT [PK_Assunto] PRIMARY KEY CLUSTERED 
(
	[CodAs] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Autor]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Autor](
	[CodAu] [int] IDENTITY(1,1) NOT NULL,
	[Nome] [varchar](40) NOT NULL,
 CONSTRAINT [PK_Autor] PRIMARY KEY CLUSTERED 
(
	[CodAu] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[LivroAutor]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LivroAutor](
	[Livro_CodAu] [int] NOT NULL,
	[Livro_CodL] [int] NOT NULL,
 CONSTRAINT [PK_LivroAutor] PRIMARY KEY CLUSTERED 
(
	[Livro_CodAu] ASC,
	[Livro_CodL] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Livro]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Livro](
	[CodL] [int] IDENTITY(1,1) NOT NULL,
	[Titulo] [varchar](40) NOT NULL,
	[Editora] [varchar](40) NOT NULL,
	[Edicao] [int] NOT NULL,
	[AnoPublicacao] [varchar](4) NOT NULL,
 CONSTRAINT [PK_Livro] PRIMARY KEY CLUSTERED 
(
	[CodL] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  View [dbo].[LivrosVisao]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[LivrosVisao] AS
select Codl Codigo,Titulo,Editora,Edicao,AnoPublicacao, Descricao as Assunto,Nome  as Autor from Livro 
left join LivroAssunto on 
LivroAssunto.Livro_Codl = Livro.CodL 
left join Assunto on 
Assunto.CodAs = LivroAssunto.Assunto_CodAs
left join LivroAutor on 
LivroAutor.Livro_CodL = Livro.CodL 
left join Autor on 
LivroAutor.Livro_CodAu = Autor.CodAu  
GO
/****** Object:  View [dbo].[AutorVisao]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[AutorVisao] AS
select Autor.CodAu as Codigo, Autor.Nome as Autor  from Autor 
GO
/****** Object:  View [dbo].[LivrosListaAutoresVisao]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[LivrosListaAutoresVisao]
	
AS

	
	WITH CodLivrosAutores AS
(
	select LivroAutor.Livro_CodL CodLivro from Autor
	left join LivroAutor on 
	Autor.CodAu = LivroAutor.Livro_CodAu
	where LivroAutor.Livro_CodL in (select Codigo from LivrosVisao 
	group by Codigo
	HAVING COUNT(Codigo) > 1)
)
	
	select Codigo,Titulo,Editora,Edicao,AnoPublicacao, Assunto,STRING_AGG(trim(Autor), ', ') Autor from LivrosVisao 
	where Codigo in (select CodLivro from CodLivrosAutores
	group by CodLivro)
	group by Codigo,Titulo,Editora,Edicao,AnoPublicacao, Assunto
	union all
	select Codigo,Titulo,Editora,Edicao,AnoPublicacao, Assunto,trim(Autor) from LivrosVisao 
	where Codigo in (select Codigo from LivrosVisao 
	group by Codigo
	HAVING COUNT(Codigo) = 1)
GO
/****** Object:  View [dbo].[AssuntoVisao]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[AssuntoVisao] AS
select Assunto.CodAs as Codigo,Descricao as Assunto  from Assunto 
GO
ALTER TABLE [dbo].[LivroAssunto]  WITH CHECK ADD  CONSTRAINT [FK_LivroAssunto_Assunto] FOREIGN KEY([Assunto_CodAs])
REFERENCES [dbo].[Assunto] ([CodAs])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[LivroAssunto] CHECK CONSTRAINT [FK_LivroAssunto_Assunto]
GO
ALTER TABLE [dbo].[LivroAssunto]  WITH CHECK ADD  CONSTRAINT [FK_LivroAssunto_Livro] FOREIGN KEY([Livro_CodL])
REFERENCES [dbo].[Livro] ([CodL])
GO
ALTER TABLE [dbo].[LivroAssunto] CHECK CONSTRAINT [FK_LivroAssunto_Livro]
GO
ALTER TABLE [dbo].[LivroAutor]  WITH CHECK ADD  CONSTRAINT [FK_LivroAutor_Autor1] FOREIGN KEY([Livro_CodAu])
REFERENCES [dbo].[Autor] ([CodAu])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[LivroAutor] CHECK CONSTRAINT [FK_LivroAutor_Autor1]
GO
ALTER TABLE [dbo].[LivroAutor]  WITH CHECK ADD  CONSTRAINT [FK_LivroAutor_Livro] FOREIGN KEY([Livro_CodL])
REFERENCES [dbo].[Livro] ([CodL])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[LivroAutor] CHECK CONSTRAINT [FK_LivroAutor_Livro]
GO
/****** Object:  StoredProcedure [dbo].[CriarAssunto]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[CriarAssunto]
	@Assunto VARCHAR(40)
AS
BEGIN
	
	insert into Assunto values (@Assunto);
	
END
GO
/****** Object:  StoredProcedure [dbo].[CriarAutor]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[CriarAutor]
	@NomeAutor VARCHAR(40)
AS
BEGIN
	
	insert into Autor values (@NomeAutor);
	
END
GO
/****** Object:  StoredProcedure [dbo].[CriarLivro]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[CriarLivro]
	@Titulo VARCHAR(40),
	@Editora VARCHAR(40),
	@Edicao int,
	@AnoPublicacao VARCHAR(4),
	@CodLNovo INT OUTPUT
AS
BEGIN
	INSERT INTO Livro 
	values (@Titulo,@Editora,@Edicao,@AnoPublicacao)
	SELECT @CodLNovo = SCOPE_IDENTITY()
END
GO
/****** Object:  StoredProcedure [dbo].[DeletarAssunto]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[DeletarAssunto]
	@CodAssunto INT
AS
BEGIN
	delete from Assunto where CodAs = @CodAssunto
END
GO
/****** Object:  StoredProcedure [dbo].[DeletarAutor]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[DeletarAutor]
	@CodAutor INT
AS
BEGIN
	delete from Autor where CodAu = @CodAutor
END
GO
/****** Object:  StoredProcedure [dbo].[DeletarLivro]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[DeletarLivro]
	@CodLivro INT
AS
BEGIN
	delete from LivroAssunto where Livro_CodL = @CodLivro 
	delete from LivroAutor where Livro_CodL = @CodLivro 
	delete from Livro where CodL = @CodLivro 
END
GO
/****** Object:  StoredProcedure [dbo].[InciarAutorAssunto]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[InciarAutorAssunto]
AS
BEGIN
	INSERT INTO Autor values ('Desconhecido') 
	INSERT INTO Assunto values ('Desconhecido') 
END
GO
/****** Object:  StoredProcedure [dbo].[LivroAssuntoAssociar]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[LivroAssuntoAssociar]
	@CodAssunto INT,
	@CodLivro INT
AS
BEGIN
	INSERT INTO LivroAssunto(Assunto_CodAs,Livro_CodL) values (@CodAssunto,@CodLivro)
END
GO
/****** Object:  StoredProcedure [dbo].[LivroAutorAssociar]    Script Date: 29/11/2023 21:03:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[LivroAutorAssociar]
	@CodAutor INT,
	@CodLivro INT
	
AS
BEGIN
	INSERT INTO LivroAutor(Livro_CodAu,Livro_CodL) values (@CodAutor,@CodLivro)
END
GO
