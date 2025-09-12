// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Hosting;

public interface IAggregable
{
    bool SupportsAggregation { get; }
    nint Wrapper { get; set; }
    IReadOnlyList<Type> AggregableInterfaces { get; }
}
