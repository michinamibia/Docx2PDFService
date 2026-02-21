# ── Build stage ───────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY Docx2PDFService.csproj .
RUN dotnet restore

COPY . .
RUN dotnet publish -c Release -o /app/publish --no-restore

# ── Runtime stage ─────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# Install LibreOffice and required fonts
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        fonts-liberation \
        fonts-dejavu \
 && apt-get clean \
 && rm -rf /var/lib/apt/lists/*

# Set the LibreOffice executable path explicitly (already on PATH in Debian)
ENV LibreOffice__ExecutablePath=libreoffice

COPY --from=build /app/publish .

EXPOSE 5525

ENTRYPOINT ["dotnet", "Docx2PDFService.dll"]
