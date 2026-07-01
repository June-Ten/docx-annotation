# syntax=docker/dockerfile:1

FROM maven:3.9-eclipse-temurin-17 AS build
WORKDIR /app

COPY pom.xml .
COPY lib/ lib/
COPY src/ src/

RUN mvn -B package -DskipTests

FROM eclipse-temurin:17-jre-jammy
WORKDIR /app

RUN mkdir -p uploads

COPY --from=build /app/target/backend-java-aspose-*.jar app.jar

EXPOSE 8080

VOLUME ["/app/uploads"]

ENTRYPOINT ["java", "-jar", "app.jar"]
