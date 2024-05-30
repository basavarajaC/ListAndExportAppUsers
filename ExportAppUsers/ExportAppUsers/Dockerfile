FROM mcr.microsoft.com/dotnet/runtime:8.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["ExportAppUsers/ExportAppUsers.csproj", "ExportAppUsers/"]
RUN dotnet restore "ExportAppUsers/ExportAppUsers.csproj"
COPY . .
WORKDIR "/src/ExportAppUsers"
RUN dotnet build "ExportAppUsers.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "ExportAppUsers.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "ExportAppUsers.dll"]
