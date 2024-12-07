package com.bb.bdd.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.servlet.FilterRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.Ordered;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;
import org.springframework.web.filter.CorsFilter;

import java.util.List;

@Configuration
public class CorsConfig {

    @Value("${cors.allowed.origins}")
    private List<String> allowedOrigins;
    @Value("${cors.allowed.headers}")
    private List<String> allowedHeaders;

    @Bean
    public FilterRegistrationBean<CorsFilter> corsFilterFilter() {
        CorsConfiguration config = new CorsConfiguration();
        config.setAllowCredentials(false);
        allowedOrigins.forEach(config::addAllowedOrigin);
        allowedHeaders.forEach(config::addAllowedHeader);
        config.addAllowedMethod("*");
        config.setMaxAge(3600L);

        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", config);

        FilterRegistrationBean<CorsFilter> filterRegistrationBean = new FilterRegistrationBean<>(new CorsFilter(source));
        filterRegistrationBean.setOrder(Ordered.HIGHEST_PRECEDENCE);

        return filterRegistrationBean;
    }
}
