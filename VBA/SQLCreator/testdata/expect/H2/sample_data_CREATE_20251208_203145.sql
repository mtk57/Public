DROP TABLE IF EXISTS sample_data;

CREATE TABLE sample_data
(
    id BIGINT NOT NULL,
    seq DECIMAL(5,0) NOT NULL,
    name VARCHAR(50),
    CONSTRAINT PK_sample_data PRIMARY KEY (id, seq)
);