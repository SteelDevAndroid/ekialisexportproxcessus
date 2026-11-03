package fr.ekialis.processus;

import java.nio.file.Path;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.core.env.Environment;

@SpringBootApplication
public class ProcessusApplication {
    public static void main(String[] args) {
        SpringApplication application = new SpringApplication(ProcessusApplication.class);
        application.setWebApplicationType(WebApplicationType.NONE);
        application.run(args).close();
    }

    @Bean
    CommandLineRunner exportProcessus(Environment environment) {
        return args -> {
            Path input = Path.of(environment.getProperty("input", "process.xlsx"));
            Path output = Path.of(environment.getProperty("output", "process-pr.xlsx"));
            int count = new ProcessusExcelExporter().export(input, output);
            System.out.printf("%d processus traités : %s%n", count, output.toAbsolutePath());
        };
    }
}
