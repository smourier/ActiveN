﻿// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Hosting;

public enum ComRegistrationTarget
{
    Default, // uses current token
    HKEY_CURRENT_USER, // always HKEY_CURRENT_USER, never requires admin but works under admin too
    HKEY_LOCAL_MACHINE, // always HKEY_LOCAL_MACHINE, may require admin
}
