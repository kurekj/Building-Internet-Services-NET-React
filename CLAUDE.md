# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Context

- Teaching repository for the **Building Internet Services** course (SGGW, winter 2026/27):
  ASP.NET Core (.NET 10) Web API + React. Public on GitHub
  (`origin` = `https://github.com/kurekj/Building-Internet-Services-NET-React`).
- The course follows TryCatchLearn **Reactivities**
  (https://github.com/TryCatchLearn/Reactivities). A local reference copy is in
  `../Reactivities/` (`API`, `Application`, `Domain`, `Infrastructure`, `Persistence`,
  `client`). Treat it as read-only reference. This repo is being built up step by step
  toward that layout, so add projects to `Reactivities.sln` as they are created.
- The parent folder (`../`) is a non-git teaching-administration folder containing student
  rosters. **Never copy roster data (raport.csv, studenci_*.csv, emaile_*.txt) into this
  repo.**

## Structure

- `Reactivities.sln`: classic `.sln` format (created with `--format sln`, not `.slnx`).
- `API/`: ASP.NET Core Web API using **controllers** (`AddControllers` / `MapControllers`,
  not minimal APIs), with built-in OpenAPI (`Microsoft.AspNetCore.OpenApi`, document at
  `/openapi/v1.json` in Development) and Swagger UI on top of it
  (`Swashbuckle.AspNetCore.SwaggerUI`, at `/swagger` in Development). It currently holds only the template
  `WeatherForecastController`.
- `Domain/`: class library for entities. It has no dependencies and holds the `Activity`
  entity (`Activity.cs`).
- `Application/`: class library for business logic and use cases. It currently holds only
  the template `Class1.cs`.
- `Persistence/`: data access with EF Core + PostgreSQL (`Npgsql.EntityFrameworkCore.PostgreSQL`):
  `AppDbContext` (`DbSet<Activity> Activities`, DateTime values stored as UTC),
  `DbInitializer` (seeds 10 sample activities if the table is empty) and `Migrations/`.
- Project references: API → Application → Persistence → Domain.
- On startup `API/Program.cs` applies pending migrations (`MigrateAsync`) and seeds data.

## Database

- PostgreSQL 16 runs in WSL (Ubuntu 24.04), database `reactivities`, login `admin`/`secret`
  (local dev only). WSL uses bridged networking, so the connection string in
  `API/appsettings.Development.json` points to the WSL IP (`192.168.7.227`), not `localhost`
  (`localhost:5432` on Windows is a separate Windows PostgreSQL). If the IP changes, update the
  connection string and the `pg_hba.conf` rule for the Windows host IP.
- Migrations: `dotnet ef migrations add <Name> -p Persistence -s API`,
  `dotnet ef database update -p Persistence -s API`.

## Commands

Run from the repo root:

```bash
dotnet build Reactivities.sln
dotnet run --project API --launch-profile http    # http://localhost:5057
dotnet watch --project API                         # hot reload during development
dotnet sln Reactivities.sln add <Path>/<Proj>.csproj
```

There are no test projects yet.

## Environment

- Windows 11, .NET SDK 10.0.201 (9.0 also installed), Node.js installed for the future
  React client. The default shell is PowerShell. The parent path contains `!`, so quote
  paths in Bash.
