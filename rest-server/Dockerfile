FROM maven:3.9.4-eclipse-temurin-17 AS build

WORKDIR /app

COPY pom.xml .
COPY src ./src

RUN mvn clean install -DskipTests

FROM openjdk:17-slim AS builder

WORKDIR /app

RUN apt-get update && apt-get install -y maven

COPY .mvn/ .mvn/
COPY mvnw pom.xml ./
RUN chmod +x mvnw

COPY src ./src
RUN ./mvnw clean package -B -DskipTests

FROM openjdk:17-slim
WORKDIR /app
COPY --from=builder /app/target/*.jar app.jar
EXPOSE 9093
ENTRYPOINT ["java", "-jar", "app.jar"]
