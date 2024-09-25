package com.etake.turnoverplan.config;

import com.etake.turnoverplan.service.StoreCategoryService;
import lombok.RequiredArgsConstructor;
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

    @Bean
    public Job job() {
        return new JobBuilder("turnoverPlanJob", jobRepository)
                .start(testStep())
                .incrementer(new RunIdIncrementer())
                .build();
    }

    @Bean
    public Step testStep() {
        return new StepBuilder("testStep", jobRepository)
                .tasklet((contribution, chunkContext) -> {
                    storeCategoryService.getSales(2024, 9);
                    return RepeatStatus.FINISHED;
                }, platformTransactionManager)
                .build();
    }
}
