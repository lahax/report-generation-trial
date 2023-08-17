package services;

import org.springframework.stereotype.Service;
import shared.Metric;

import java.util.ArrayList;
import java.util.Map;

@Service
public class DataNormalizationService {

    public ArrayList<Metric> normalizePayload(String payload) {
        return new ArrayList<Metric>();
    }
}