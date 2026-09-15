using System;

namespace RevitMcpStdioProxy
{
    internal sealed class ProxyOptions
    {
        public string DirectEndpoint { get; private set; }
        public string InstanceId { get; private set; }
        public int? ProcessId { get; private set; }
        public int? RevitVersion { get; private set; }
        public string ModelName { get; private set; }

        public bool UsesDiscovery
        {
            get { return string.IsNullOrWhiteSpace(DirectEndpoint); }
        }

        public bool HasDiscoverySelectors
        {
            get
            {
                return !string.IsNullOrWhiteSpace(InstanceId)
                    || ProcessId.HasValue
                    || RevitVersion.HasValue
                    || !string.IsNullOrWhiteSpace(ModelName);
            }
        }

        public static ProxyOptions Parse(string[] args)
        {
            ProxyOptions options = new ProxyOptions();

            for (int i = 0; args != null && i < args.Length; i++)
            {
                string argument = args[i] ?? string.Empty;
                if (string.Equals(argument, "--endpoint", StringComparison.OrdinalIgnoreCase))
                {
                    options.DirectEndpoint = NormalizeEndpoint(ReadValue(args, ref i, argument));
                }
                else if (string.Equals(argument, "--instance-id", StringComparison.OrdinalIgnoreCase))
                {
                    options.InstanceId = ReadValue(args, ref i, argument).Trim();
                }
                else if (string.Equals(argument, "--process-id", StringComparison.OrdinalIgnoreCase))
                {
                    options.ProcessId = ParsePositiveInt(ReadValue(args, ref i, argument), argument);
                }
                else if (string.Equals(argument, "--revit-version", StringComparison.OrdinalIgnoreCase))
                {
                    options.RevitVersion = ParsePositiveInt(ReadValue(args, ref i, argument), argument);
                }
                else if (string.Equals(argument, "--model", StringComparison.OrdinalIgnoreCase))
                {
                    options.ModelName = ReadValue(args, ref i, argument).Trim();
                }
            }

            if (string.IsNullOrWhiteSpace(options.DirectEndpoint))
            {
                string environmentEndpoint = Environment.GetEnvironmentVariable("REVIT_MCP_ENDPOINT");
                if (!string.IsNullOrWhiteSpace(environmentEndpoint))
                    options.DirectEndpoint = NormalizeEndpoint(environmentEndpoint);
            }

            return options;
        }

        public string Describe()
        {
            if (!UsesDiscovery)
                return "direct endpoint " + DirectEndpoint;

            return "instance discovery"
                + DescribeValue("instance", InstanceId)
                + DescribeValue("process", ProcessId)
                + DescribeValue("version", RevitVersion)
                + DescribeValue("model", ModelName);
        }

        private static string ReadValue(string[] args, ref int index, string argument)
        {
            if (args == null || index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
                throw new ArgumentException("Missing value for " + argument + ".");

            index++;
            return args[index];
        }

        private static int ParsePositiveInt(string value, string argument)
        {
            int result;
            if (!int.TryParse(value, out result) || result <= 0)
                throw new ArgumentException(argument + " must be a positive integer.");

            return result;
        }

        private static string NormalizeEndpoint(string endpoint)
        {
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new ArgumentException("MCP endpoint cannot be empty.");

            endpoint = endpoint.Trim();
            return endpoint.EndsWith("/", StringComparison.Ordinal) ? endpoint : endpoint + "/";
        }

        private static string DescribeValue(string name, object value)
        {
            return value == null || string.IsNullOrWhiteSpace(value.ToString())
                ? string.Empty
                : ", " + name + "=" + value;
        }
    }
}
