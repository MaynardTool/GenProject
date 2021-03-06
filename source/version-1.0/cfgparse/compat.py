# Copyright (c) 2001, 2002, 2003 Python Software Foundation
# Copyright (c) 2004 Paramjit Oberoi <param.cs.wisc.edu>
# All Rights Reserved.  See LICENSE-PSF & LICENSE for details.

"""Compatibility interfaces for ConfigParser

Interfaces of ConfigParser, RawConfigParser and SafeConfigParser
should be completely identical to the Python standard library
versions.  Tested with the unit tests included with Python-2.3.4

The underlying ini_namespace object can be accessed as cfg.data
"""

import re
from ConfigParser import DuplicateSectionError,    \
                  NoSectionError, NoOptionError,   \
                  InterpolationMissingOptionError, \
                  InterpolationDepthError,         \
                  InterpolationSyntaxError,        \
                  DEFAULTSECT, MAX_INTERPOLATION_DEPTH

import iniparser

class RawConfigParser(object):
    def __init__(self, defaults=None):
        self.data = iniparser.ini_namespace(defaults=defaults)

    def defaults(self):
        d = {}
        for name, lineobj in self.data._defaults._options:
            d[name] = lineobj.value
        return d

    def sections(self):
        """Return a list of section names, excluding [DEFAULT]"""
        return list(self.data)

    def add_section(self, section):
        """Create a new section in the configuration.

        Raise DuplicateSectionError if a section by the specified name
        already exists.
        """
        if self.has_section(section):
            raise DuplicateSectionError(section)
        else:
            self.data.new_namespace(section)

    def has_section(self, section):
        """Indicate whether the named section is present in the configuration.

        The DEFAULT section is not acknowledged.
        """
        try:
            self.data[section]
            return True
        except KeyError:
            return False

    def options(self, section):
        """Return a list of option names for the given section name."""
        try:
            return list(self.data[section])
        except KeyError:
            raise NoSectionError(section)

    def read(self, filenames):
        """Read and parse a filename or a list of filenames.

        Files that cannot be opened are silently ignored; this is
        designed so that you can specify a list of potential
        configuration file locations (e.g. current directory, user's
        home directory, systemwide directory), and all existing
        configuration files in the list will be read.  A single
        filename may also be given.
        """
        if isinstance(filenames, basestring):
            filenames = [filenames]
        for filename in filenames:
            try:
                fp = open(filename)
            except IOError:
                continue
            self.data.readfp(fp)
            fp.close()

    def readfp(self, fp, filename=None):
        """Like read() but the argument must be a file-like object.

        The `fp' argument must have a `readline' method.  Optional
        second argument is the `filename', which if not given, is
        taken from fp.name.  If fp has no `name' attribute, `<???>' is
        used.
        """
        self.data.readfp(fp)

    def get(self, section, option, vars=None):
        if not self.has_section(section):
            raise NoSectionError(section)
        if vars is not None and option in vars:
            value = vars[option]
        try:
            return self.data[section][option]
        except KeyError:
            raise NoOptionError(option, section)

    def items(self, section):
        try:
            ans = []
            for opt in self.data[section]:
                ans.append((opt, self.data[section][opt]))
            return ans
        except KeyError:
            raise NoSectionError(section)

    def getint(self, section, option):
        return int(self.get(section, option))

    def getfloat(self, section, option):
        return float(self.get(section, option))

    _boolean_states = {'1': True, 'yes': True, 'true': True, 'on': True,
                       '0': False, 'no': False, 'false': False, 'off': False}

    def getboolean(self, section, option):
        v = self.get(section, option)
        if v.lower() not in self._boolean_states:
            raise ValueError, 'Not a boolean: %s' % v
        return self._boolean_states[v.lower()]

    def get_oxf(self): return self.data._optionxform
    def set_oxf(self, value): self.data._optionxform = value
    optionxform = property(get_oxf, set_oxf)

    def has_option(self, section, option):
        """Check for the existence of a given option in a given section."""
        try:
            sec = self.data[section]
        except KeyError:
            raise NoSectionError(section)
        try:
            sec[option]
            return True
        except KeyError:
            return False

    def set(self, section, option, value):
        """Set an option."""
        try:
            self.data[section][option] = value
        except KeyError:
            raise NoSectionError(section)

    def write(self, fp):
        """Write an .ini-format representation of the configuration state."""
        fp.write(str(self.data))

    def remove_option(self, section, option):
        """Remove an option."""
        try:
            sec = self.data[section]
        except KeyError:
            raise NoSectionError(section)
        try:
            sec[option]
            del sec[option]
            return 1
        except KeyError:
            return 0

    def remove_section(self, section):
        """Remove a file section."""
        if not self.has_section(section):
            return False
        del self.data[section]
        return True


class cfg_dict(object):
    def __init__(self, cfg, section, vars):
        self.cfg = cfg
        self.section = section
        self.vars = vars

    def __getitem__(self, key):
        try:
            return RawConfigParser.get(self.cfg, self.section, key, self.vars)
        except (NoOptionError, NoSectionError):
            raise KeyError(key)


class ConfigParser(RawConfigParser):

    def get(self, section, option, raw=False, vars=None):
        """Get an option value for a given section.

        All % interpolations are expanded in the return values, based on the
        defaults passed into the constructor, unless the optional argument
        `raw' is true.  Additional substitutions may be provided using the
        `vars' argument, which must be a dictionary whose contents overrides
        any pre-existing defaults.

        The section DEFAULT is special.
        """
        if section != DEFAULTSECT and not self.has_section(section):
            raise NoSectionError(section)

        option = self.optionxform(option)
        value = RawConfigParser.get(self, section, option, vars)

        if raw:
            return value
        else:
            d = cfg_dict(self, section, vars)
            return self._interpolate(section, option, value, d)

    def _interpolate(self, section, option, rawval, vars):
        # do the string interpolation
        value = rawval
        depth = MAX_INTERPOLATION_DEPTH
        while depth:                    # Loop through this until it's done
            depth -= 1
            if value.find("%(") != -1:
                try:
                    value = value % vars
                except KeyError, e:
                    raise InterpolationMissingOptionError(
                        option, section, rawval, e[0])
            else:
                break
        if value.find("%(") != -1:
            raise InterpolationDepthError(option, section, rawval)
        return value

    def items(self, section, raw=False, vars=None):
        """Return a list of tuples with (name, value) for each option
        in the section.

        All % interpolations are expanded in the return values, based on the
        defaults passed into the constructor, unless the optional argument
        `raw' is true.  Additional substitutions may be provided using the
        `vars' argument, which must be a dictionary whose contents overrides
        any pre-existing defaults.

        The section DEFAULT is special.
        """
        if section != DEFAULTSECT and not self.has_section(section):
            raise NoSectionError(section)
        if vars is None:
            options = list(self.data[section])
        else:
            options = []
            for x in self.data[section]:
                if x not in vars:
                    options.append(x)
            options.extend(vars.keys())

        if "__name__" in options:
            options.remove("__name__")

        d = cfg_dict(self, section, vars)
        if raw:
            return [(option, d[option])
                    for option in options]
        else:
            return [(option, self._interpolate(section, option, d[option], d))
                    for option in options]


class SafeConfigParser(ConfigParser):

    def _interpolate(self, section, option, rawval, vars):
        # do the string interpolation
        L = []
        self._interpolate_some(option, L, rawval, section, vars, 1)
        return ''.join(L)

    _interpvar_match = re.compile(r"%\(([^)]+)\)s").match

    def _interpolate_some(self, option, accum, rest, section, map, depth):
        if depth > MAX_INTERPOLATION_DEPTH:
            raise InterpolationDepthError(option, section, rest)
        while rest:
            p = rest.find("%")
            if p < 0:
                accum.append(rest)
                return
            if p > 0:
                accum.append(rest[:p])
                rest = rest[p:]
            # p is no longer used
            c = rest[1:2]
            if c == "%":
                accum.append("%")
                rest = rest[2:]
            elif c == "(":
                m = self._interpvar_match(rest)
                if m is None:
                    raise InterpolationSyntaxError(option, section,
                        "bad interpolation variable reference %r" % rest)
                var = m.group(1)
                rest = rest[m.end():]
                try:
                    v = map[var]
                except KeyError:
                    raise InterpolationMissingOptionError(
                        option, section, rest, var)
                if "%" in v:
                    self._interpolate_some(option, accum, v,
                                           section, map, depth + 1)
                else:
                    accum.append(v)
            else:
                raise InterpolationSyntaxError(
                    option, section,
                    "'%' must be followed by '%' or '(', found: " + `rest`)
