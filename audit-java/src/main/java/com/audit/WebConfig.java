package com.audit;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

// 跨域配置：允许 Office Online 等外部服务通过浏览器访问 /open/ 接口
@Configuration
public class WebConfig implements WebMvcConfigurer {
    @Override
    public void addCorsMappings(CorsRegistry registry) {
        registry.addMapping("/**")           // 所有路径
            .allowedOrigins("*")             // 允许所有来源
            .allowedMethods("GET", "POST", "OPTIONS")
            .allowedHeaders("*");
    }
}
