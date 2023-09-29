#!/usr/bin/env python
from ecmwfapi import ECMWFDataServer
server = ECMWFDataServer()
server.retrieve({
    "class": "s2",
    "dataset": "s2s",
    "date": "2023-01-02/2023-01-05/2023-01-09/2023-01-12/2023-01-16/2023-01-19/2023-01-23/2023-01-26/2023-01-30",
    "expver": "prod",
    "levtype": "sfc",
    "model": "glob",
    "number": "1/2/3/4/5/6/7/8/9/10/11/12/13/14/15/16/17/18/19/20/21/22/23/24/25/26/27/28/29/30/31/32/33/34/35/36/37/38/39/40/41/42/43/44/45/46/47/48/49/50",
    "origin": "ecmf",
    "param": "228228",
    "step": "0/6",
    "stream": "enfo",
    "time": "00:00:00",
    "type": "pf",
    "target": "output",
})
