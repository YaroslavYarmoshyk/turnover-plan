package com.etake.turnoverplan.config;

import com.etake.turnoverplan.config.properties.SystemConfigurationProperties;
import com.etake.turnoverplan.service.ExcelService;
import com.etake.turnoverplan.service.StoreCategoryService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.batch.core.Job;
import org.springframework.batch.core.Step;
import org.springframework.batch.core.job.builder.JobBuilder;
import org.springframework.batch.core.launch.support.RunIdIncrementer;
import org.springframework.batch.core.repository.JobRepository;
import org.springframework.batch.core.step.builder.StepBuilder;
import org.springframework.batch.repeat.RepeatStatus;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;
import org.springframework.transaction.PlatformTransactionManager;

@Component
@RequiredArgsConstructor
public class BatchJobConfig {
    private final JobRepository jobRepository;
    private final PlatformTransactionManager platformTransactionManager;

    private final StoreCategoryService storeCategoryService;
    private final ExcelService excelService;

    private final SystemConfigurationProperties systemConfigurationProperties;

    @Bean
    public Job job() {
        return new JobBuilder("turnoverPlanJob", jobRepository)
//                .start(testStep())
                .start(writeStep())
                .incrementer(new RunIdIncrementer())
                .build();
    }

//    @Bean
//    public Step testStep() {
//        return new StepBuilder("testStep", jobRepository)
//                .tasklet((contribution, chunkContext) -> {
//                    storeCategoryService.getSales(2024, 9);
//                    return RepeatStatus.FINISHED;
//                }, platformTransactionManager)
//                .build();
//    }

    @Bean
    public Step writeStep() {
        return new StepBuilder("writeStep", jobRepository)
                .tasklet((contribution, chunkContext) -> {
                    final Workbook workbook = excelService.getWorkbook(storeCategoryService.getSales());
//                    final Workbook workbook = excelService.getWorkbook(getTestSales());
                    excelService.writeWorkbook(workbook, systemConfigurationProperties.filePath() + "test.xlsx");
                    return RepeatStatus.FINISHED;
                }, platformTransactionManager)
                .build();
    }
}
