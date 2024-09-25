package com.etake.turnoverplan.config.properties;

import org.springframework.boot.context.properties.ConfigurationProperties;


@ConfigurationProperties(prefix = "system-configuration")
public record SystemConfigurationProperties(Integer year, Integer month) {
}
