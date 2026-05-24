/**
 * ============================================
 * JOB BOARD API + ADMIN BACKEND
 * ============================================
 *
 * FEATURES:
 * - Get all jobs
 * - Get single job
 * - Search jobs
 * - Filter jobs
 * - Get statistics
 * - Get filters
 * - Add jobs from admin panel
 * - Featured jobs
 * - Secret key protection
 *
 * ============================================
 * GOOGLE SHEET STRUCTURE
 * ============================================
 *
 * A  Job ID
 * B  Job Title
 * C  Company Name
 * D  Company Logo URL
 * E  Location
 * F  Job Type
 * G  Category
 * H  Experience Level
 * I  Salary Range
 * J  Job Description
 * K  Responsibilities
 * L  Requirements
 * M  Skills Required
 * N  Benefits
 * O  Apply URL
 * P  Posted Date
 * Q  Application Deadline
 * R  Status
 * S  Featured
 *
 */

const SHEET_NAME = 'Jobs';

/**
 * ============================================
 * CHANGE THIS SECRET KEY
 * ============================================
 */

const ADMIN_SECRET = 'CHANGE_THIS_SECRET_KEY';

/**
 * ============================================
 * GET REQUESTS
 * ============================================
 */

function doGet(e) {

  const params = (e && e.parameter) || {};
  const action = params.action || 'getJobs';

  try {

    let payload;

    if (action === 'getStats') {

      payload = {
        success: true,
        data: getStats_()
      };

    } else if (action === 'getFilters') {

      payload = {
        success: true,
        data: getFilters_()
      };

    } else if (action === 'getJob') {

      payload = {
        success: true,
        data: getJob_(params.id || '')
      };

    } else if (action === 'search') {

      const jobs = searchJobs_(
        getActiveJobs_(),
        getParam_(params, 'q') || ''
      );

      payload = {
        success: true,
        data: jobs,
        count: jobs.length
      };

    } else if (action === 'filter') {

      const jobs = filterJobs_(
        getActiveJobs_(),
        params
      );

      payload = {
        success: true,
        data: jobs,
        count: jobs.length
      };

    } else if (action === 'getJobs') {

      let jobs = getActiveJobs_();

      if (
        String(params.featured).toLowerCase() === 'true'
      ) {
        jobs = jobs.filter(job => job.featured);
      }

      payload = {
        success: true,
        data: jobs,
        count: jobs.length
      };

    } else {

      throw new Error(
        'Unsupported action: ' + action
      );

    }

    return jsonResponse_(
      payload,
      params.callback
    );

  } catch(error) {

    return jsonResponse_({
      success: false,
      error: error.message || String(error)
    });

  }

}

/**
 * ============================================
 * POST REQUESTS (ADMIN PANEL)
 * ============================================
 */

function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received');
    }

    const data = JSON.parse(e.postData.contents);

    /**
     * ============================================
     * SECURITY CHECK
     * ============================================
     */

    if (
      !data.secret ||
      data.secret !== ADMIN_SECRET
    ) {
      throw new Error('Unauthorized access');
    }

    const sheet = getSheet_();

    sheet.appendRow([

      data.jobId || '',
      data.jobTitle || '',
      data.companyName || '',
      data.logoUrl || '',
      data.location || '',
      data.jobType || '',
      data.category || '',
      data.experience || '',
      data.salary || '',
      data.description || '',
      data.responsibilities || '',
      data.requirements || '',
      data.skills || '',
      data.benefits || '',
      data.applyUrl || '',
      data.postedDate || '',
      data.deadline || '',
      data.status || 'Active',
      data.featured || 'FALSE'

    ]);

    return jsonResponse_({
      success: true,
      message: 'Job added successfully'
    });

  } catch(error) {

    return jsonResponse_({
      success: false,
      message: error.message || String(error)
    });

  }

}

/**
 * ============================================
 * OPTIONS REQUEST
 * ============================================
 */

function doOptions() {

  return jsonResponse_({
    success: true,
    data: {}
  });

}

/**
 * ============================================
 * JSON RESPONSE
 * ============================================
 */

function jsonResponse_(payload, callback) {

  const body = callback
    ? String(callback)
        .replace(/[^\w$.]/g, '') +
        '(' +
        JSON.stringify(payload) +
        ');'
    : JSON.stringify(payload);

  const output = ContentService
    .createTextOutput(body)
    .setMimeType(
      callback
        ? ContentService.MimeType.JAVASCRIPT
        : ContentService.MimeType.JSON
    );

  return output;

}

/**
 * ============================================
 * GET SHEET
 * ============================================
 */

function getSheet_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet named "Jobs" was not found.'
    );
  }

  return sheet;

}

/**
 * ============================================
 * GET ROWS
 * ============================================
 */

function getRows_() {

  const values = getSheet_()
    .getDataRange()
    .getValues();

  if (values.length < 2) {
    return [];
  }

  return values
    .slice(1)
    .filter(row =>
      row.some(cell =>
        cell !== '' &&
        cell !== null
      )
    )
    .map(rowToJob_);

}

/**
 * ============================================
 * GET ACTIVE JOBS
 * ============================================
 */

function getActiveJobs_() {

  return sortJobs_(
    getRows_().filter(
      job =>
        normalize_(job.status) === 'active'
    )
  );

}

/**
 * ============================================
 * ROW TO JOB OBJECT
 * ============================================
 */

function rowToJob_(row) {

  return {

    id: clean_(row[0]),
    title: clean_(row[1]),
    company: clean_(row[2]),
    logoUrl: clean_(row[3]),
    location: clean_(row[4]),
    type: clean_(row[5]),
    category: clean_(row[6]),
    experience: clean_(row[7]),
    salary: clean_(row[8]),

    description: clean_(row[9]),

    responsibilities: splitList_(
      row[10],
      '|'
    ),

    requirements: splitList_(
      row[11],
      '|'
    ),

    skills: splitList_(
      row[12],
      ','
    ),

    benefits: splitList_(
      row[13],
      '|'
    ),

    applyUrl: clean_(row[14]),

    postedDate: formatDate_(row[15]),

    deadline: formatDate_(row[16]),

    status: clean_(row[17]),

    featured: parseBoolean_(row[18])

  };

}

/**
 * ============================================
 * GET SINGLE JOB
 * ============================================
 */

function getJob_(id) {

  const needle = normalize_(id);

  if (!needle) {
    return null;
  }

  return getActiveJobs_().find(
    job =>
      normalize_(job.id) === needle
  ) || null;

}

/**
 * ============================================
 * SEARCH JOBS
 * ============================================
 */

function searchJobs_(jobs, query) {

  const needle = normalize_(query);

  if (!needle) {
    return jobs;
  }

  return jobs.filter(job => [

    job.title,
    job.company,
    job.location,
    job.skills.join(' ')

  ].some(value =>

    normalize_(value)
      .indexOf(needle) !== -1

  ));

}

/**
 * ============================================
 * FILTER JOBS
 * ============================================
 */

function filterJobs_(jobs, params) {

  let filtered = jobs.slice();

  filtered = filterByCsv_(
    filtered,
    getParam_(params, 'category'),
    'category'
  );

  filtered = filterByCsv_(
    filtered,
    getParam_(params, 'type'),
    'type'
  );

  filtered = filterByCsv_(
    filtered,
    getParam_(params, 'experience'),
    'experience'
  );

  filtered = filterByCsv_(
    filtered,
    getParam_(params, 'location'),
    'location'
  );

  filtered = filterByCsv_(
    filtered,
    getParam_(params, 'company'),
    'company'
  );

  if (
    String(
      getParam_(params, 'featured')
    ).toLowerCase() === 'true'
  ) {

    filtered = filtered.filter(
      job => job.featured
    );

  }

  const dateRange = getParam_(
    params,
    'dateRange'
  );

  if (
    dateRange &&
    dateRange !== 'any'
  ) {

    const now = new Date();

    const days =
      dateRange === '24h'
        ? 1
        : Number(
            String(dateRange)
              .replace(/\D/g, '')
          );

    if (days > 0) {

      const cutoff = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate() - days
      );

      filtered = filtered.filter(job => {

        const posted = parseDate_(
          job.postedDate
        );

        return posted && posted >= cutoff;

      });

    }

  }

  if (getParam_(params, 'q')) {

    filtered = searchJobs_(
      filtered,
      getParam_(params, 'q')
    );

  }

  return sortJobs_(filtered);

}

/**
 * ============================================
 * GET PARAM
 * ============================================
 */

function getParam_(params, key) {

  if (params[key]) {
    return params[key];
  }

  if (
    params.parameters &&
    params.parameters[key]
  ) {
    return params.parameters[key].join(',');
  }

  return '';

}

/**
 * ============================================
 * FILTER CSV
 * ============================================
 */

function filterByCsv_(jobs, rawValue, key) {

  const values = splitList_(
    rawValue,
    ','
  ).map(normalize_);

  if (!values.length) {
    return jobs;
  }

  return jobs.filter(job =>

    values.indexOf(
      normalize_(job[key])
    ) !== -1

  );

}

/**
 * ============================================
 * GET STATS
 * ============================================
 */

function getStats_() {

  const jobs = getActiveJobs_();

  const companies = unique_(
    jobs.map(job => job.company)
  );

  const categories = unique_(
    jobs.map(job => job.category)
  );

  const weekAgo = new Date();

  weekAgo.setDate(
    weekAgo.getDate() - 7
  );

  return {

    totalJobs: jobs.length,

    totalCompanies: companies.length,

    categories: categories,

    newThisWeek: jobs.filter(job => {

      const posted = parseDate_(
        job.postedDate
      );

      return posted && posted >= weekAgo;

    }).length

  };

}

/**
 * ============================================
 * GET FILTERS
 * ============================================
 */

function getFilters_() {

  const jobs = getActiveJobs_();

  return {

    categories: unique_(
      jobs.map(job => job.category)
    ),

    types: unique_(
      jobs.map(job => job.type)
    ),

    experiences: unique_(
      jobs.map(job => job.experience)
    ),

    locations: unique_(
      jobs.map(job => job.location)
    ),

    companies: unique_(
      jobs.map(job => job.company)
    )

  };

}

/**
 * ============================================
 * SORT JOBS
 * ============================================
 */

function sortJobs_(jobs) {

  return jobs.slice().sort((a, b) => {

    if (
      a.featured !== b.featured
    ) {
      return a.featured ? -1 : 1;
    }

    const dateA = parseDate_(
      a.postedDate
    );

    const dateB = parseDate_(
      b.postedDate
    );

    return (
      (dateB
        ? dateB.getTime()
        : 0) -
      (dateA
        ? dateA.getTime()
        : 0)
    );

  });

}

/**
 * ============================================
 * UNIQUE VALUES
 * ============================================
 */

function unique_(values) {

  const seen = {};

  return values

    .map(clean_)

    .filter(Boolean)

    .filter(value => {

      const key = normalize_(value);

      if (seen[key]) {
        return false;
      }

      seen[key] = true;

      return true;

    })

    .sort((a, b) =>
      a.localeCompare(b)
    );

}

/**
 * ============================================
 * SPLIT LIST
 * ============================================
 */

function splitList_(value, delimiter) {

  return clean_(value)

    .split(delimiter)

    .map(item => item.trim())

    .filter(Boolean);

}

/**
 * ============================================
 * CLEAN VALUE
 * ============================================
 */

function clean_(value) {

  if (value instanceof Date) {
    return formatDate_(value);
  }

  return (
    value === null ||
    value === undefined
  )
    ? ''
    : String(value).trim();

}

/**
 * ============================================
 * NORMALIZE
 * ============================================
 */

function normalize_(value) {

  return clean_(value)
    .toLowerCase();

}

/**
 * ============================================
 * BOOLEAN PARSER
 * ============================================
 */

function parseBoolean_(value) {

  return [
    'true',
    'yes',
    '1'
  ].indexOf(
    normalize_(value)
  ) !== -1;

}

/**
 * ============================================
 * FORMAT DATE
 * ============================================
 */

function formatDate_(value) {

  if (!value) {
    return '';
  }

  if (value instanceof Date) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  }

  return clean_(value);

}

/**
 * ============================================
 * PARSE DATE
 * ============================================
 */

function parseDate_(value) {

  if (!value) {
    return null;
  }

  const date = new Date(value);

  return isNaN(date.getTime())
    ? null
    : date;

}
