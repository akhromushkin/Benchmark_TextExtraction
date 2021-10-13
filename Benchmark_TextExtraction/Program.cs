using BenchmarkDotNet.Running;
using System;

namespace Benchmark_TextExtraction {
    class Program {
        static void Main(string[] args) {
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
        }
    }
}
